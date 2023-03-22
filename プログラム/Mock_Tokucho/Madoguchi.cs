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
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;
using Point = System.Drawing.Point;
using System.Collections;
using System.Text.RegularExpressions;
using System.Deployment.Application;
using C1.C1Excel;
using TokuchoBugyoK2;
using System.Threading;

namespace TokuchoBugyoK2
{
    public partial class Madoguchi : Form
    {
        public string[] UserInfos;
        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();
        private DataTable comboDt3 = new DataTable();
        public Boolean ReSearch = false;
        private Boolean init_flg = true;
        //private Boolean kensakukikanFlg = false;
        private string fileName;
        private int readCount = 0;
        private int errorId = 0;
        private int chousahinmokuErrorId = 0;
        private string errorUser = "";
        private string saibanMadoguchiID = "";
        private string ErrorMsg = "";
        private int exceptionFlg = 0; // 例外発生フラグ 0:発生なし 1:発生
        private string tokuchoBangou = "";
        private string tokuchoBangouEda = ""; // 報告書フォルダ配下の枝番フォルダ削除に使用
        private string ShukeiHyoFolder = "";
        private string HoukokuShoFolder = "";
        private string ShiryouHolder = "";

        private static int sheetFlg1 = 0; // 調査概要タブ存在フラグ 0:あり 1:なし
        private static int sheetFlg2 = 0; // 担当部所タブ存在フラグ 0:あり 1:なし
        private static int sheetFlg3 = 0; // 調査品目一覧タブ存在フラグ 0:あり 1:なし
        private static int sheetFlg4 = 0; // 協力依頼書タブ存在フラグ 0:あり 1:なし
        private static int sheetFlg5 = 0; // 応援受付状況タブ存在フラグ 0:あり 1:なし
        private static int sheetFlg6 = 0; // 単品入力項目タブ存在フラグ 0:あり 1:なし

        // WorkBook Open フラグ 0:閉じている 1:開いている
        private static int workBookOpenFlg = 0;

        // 処理待ち時間
        private static int waitTime = 3;

        private Color errorBackColor = Color.FromArgb(255, 204, 255);

        public Madoguchi()
        {
            InitializeComponent();

            // コンボボックスにマウスホイールイベントを付与
            this.item_Nendo.MouseWheel += item_MouseWheel;
            this.item_ChousaBusho.MouseWheel += item_MouseWheel;
            this.item_DateKindCombo.MouseWheel += item_MouseWheel;
            this.item_JutakuBusho.MouseWheel += item_MouseWheel;
            this.item_MadoguchiBusho.MouseWheel += item_MouseWheel;
            this.item_FromTo.MouseWheel += item_MouseWheel;
            this.item_ShimekiriSentaku.MouseWheel += item_MouseWheel;
            this.item_ChousaKind.MouseWheel += item_MouseWheel;
            this.item_Shijisho.MouseWheel += item_MouseWheel;
            this.item_JishiKbn.MouseWheel += item_MouseWheel;
            this.item_Shintyokujyoukyo.MouseWheel += item_MouseWheel;
            this.item_KanrityouhyouInsatu.MouseWheel += item_MouseWheel;
            this.item_Hyoujikensuu.MouseWheel += item_MouseWheel;

            //エントリ君修正STEP2
            this.ErrorMessage.Font = new System.Drawing.Font(this.ErrorMessage.Font.Name, float.Parse(GlobalMethod.GetCommonValue1("DSP_ERROR_FONTSIZE")));

        }

        private void Madoguchi_Load(object sender, EventArgs e)
        {
            //不具合No1355（1123）
            lblVersion.Text = GlobalMethod.GetCommonValue1("APL_VERSION");
            if (GlobalMethod.GetCommonValue1("BOOT_MODE") == "1")
            {
                lblBootMode.Text = GlobalMethod.GetCommonValue2("BOOT_MODE");
            }
            //ユーザ名を設定
            label3.Text = UserInfos[3] + "：" + UserInfos[1];

            //c1FlexGrid1[1, 9] = "0";
            //c1FlexGrid1[2, 9] = "1";

            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            Hashtable imgMap = new Hashtable();

            gridSizeChange();

            //imgMap.Add("0", Image.FromFile(@"./Resource/OnegaiIcon35px.png"));
            //imgMap.Add("1", Image.FromFile(@"./Resource/kan.png"));

            //ソート項目にアイコンを設定
            C1.Win.C1FlexGrid.CellRange cr;
            Bitmap bmp1 = new Bitmap("Resource/Image/SortIconDefalt.png");
            Bitmap bmpSort = new Bitmap(bmp1, bmp1.Width / 6, bmp1.Height / 6);
            cr = c1FlexGrid1.GetCellRange(0, 1);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;

            // 特調番号～最終検査
            for(int i = 3; i < c1FlexGrid1.Cols.Count; i++)
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
            imgMap.Add("4", Image.FromFile("Resource/Image/blank2.png"));      // 上記のいずれにも該当しない
            c1FlexGrid1.Cols[1].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[1].ImageMap = imgMap;
            c1FlexGrid1.Cols[1].ImageAndText = false;


            //c1FlexGrid1.Cols[9].ImageMap = imgMap;
            //c1FlexGrid1.Cols[9].ImageAndText = false;

            //編集の画像切り替え
            imgMap = new Hashtable();
            imgMap.Add("0", Image.FromFile("Resource/Image/file_presentation1_g.png"));
            imgMap.Add("1", Image.FromFile("Resource/Image/file_presentation1.png"));
            c1FlexGrid1.Cols[2].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[2].ImageMap = imgMap;
            c1FlexGrid1.Cols[2].ImageAndText = false;

            // 応援依頼編集の画像切り替え
            imgMap = new Hashtable();
            // 協力依頼書画面から依頼書出力はお相撲さんアイコン、受付状況画面の応援状況をチェック時は丸完アイコン
            //imgMap.Add(1, Image.FromFile("Resource/OnegaiIcon35px.png"));
            //imgMap.Add(1, Image.FromFile("Resource/kan.png"));
            imgMap.Add(1, Image.FromFile("Resource/OnegaiIcon35px.png"));
            imgMap.Add(2, Image.FromFile("Resource/kan.png"));
            c1FlexGrid1.Cols[9].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[9].ImageMap = imgMap;
            c1FlexGrid1.Cols[9].ImageAndText = false;

            // 応援完了編集の画像切り替え
            imgMap = new Hashtable();
            imgMap.Add(1, Image.FromFile("Resource/kan.png"));
            c1FlexGrid1.Cols[10].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[10].ImageMap = imgMap;
            c1FlexGrid1.Cols[10].ImageAndText = false;

            set_combo();
            // 検索条件初期化
            ClearForm();
            // 窓口部所に自分の部所を入れる
            item_MadoguchiBusho.SelectedValue = UserInfos[2];

            // 完了は下はデフォルトチェック
            item_Kanryouhasita.Checked = true;

            //一覧に表示するデータを取得
            get_data();
            init_flg = false;
            //kensakukikanFlg = false;

        }


        private void textbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        // 新規ボタン
        private void BtnInsert_Click(object sender, EventArgs e)
        {
            this.ReSearch = true;
            Madoguchi_Input form = new Madoguchi_Input();
            form.mode = "insert";
            form.UserInfos = this.UserInfos;
            form.Show(this);

            this.Hide();
        }
        // TOP
        private void button6_Click(object sender, EventArgs e)
        {
            this.ReSearch = true;
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
        // ヘッダーの窓口ボタン
        private void button7_Click(object sender, EventArgs e)
        {
            //Madoguchi form = new Madoguchi();
            //form.UserInfos = this.UserInfos;
            //form.Show();
            //this.Close();

            //Form f = null;
            //Boolean openFlg = false;
            //for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
            //{
            //    f = System.Windows.Forms.Application.OpenForms[i];
            //    if (f.Text.IndexOf("窓口ミハル") >= 0 && f.Text.IndexOf("編集") <= -1)
            //    {
            //        f.Show();
            //        openFlg = true;
            //        break;
            //    }
            //}
            //if (!openFlg)
            //{
            //    Madoguchi form = new Madoguchi();
            //    form.UserInfos = this.UserInfos;
            //    form.Show();
            //    //this.Close();
            //}
            //this.Hide();
        }
        // 特命課長
        private void button3_Click(object sender, EventArgs e)
        {
            this.ReSearch = true;
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
            this.Hide();
        }
        // 自分大臣
        private void button8_Click(object sender, EventArgs e)
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

        private void get_data(string ikkatsuFlg = "0")
        {
            string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            //データ取得処理
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();


                // 打診中の業務
                if (item_DashinGyoumu.Checked)
                {
                    // 打診中の業務にチェックし、検索した場合、業務区分を2:打診中にして検索する
                    item_JishiKbn.SelectedValue = 2;
                }

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

                // 備考1～25取得処理
                comboDt3 = new DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                    "GyoumuBushoCD " +
                    ",substring(BushokanriboKameiRaku,1,2) " +
                    ",KashoShibuCD " +
                    ",BushoShibuCD " +
                    "FROM Mst_Busho " +
                    "WHERE " +
                    "BushoMadoguchiHyoujiFlg = 1  " +
                    "AND KashoShibuCD != '' " +
                    "AND BushoDeleteFlag != 1 " +
                    "AND GyoumuBushoCD LIKE '127%' " +
                    "or GyoumuBushoCD LIKE '17%' " +
                    "or GyoumuBushoCD LIKE '16%' " +
                    "AND not GyoumuBushoCD LIKE '1279%' " +
                    "AND BushoDeleteFlag != 1 " +
                    "ORDER BY BushoMadoguchiNarabijun";

                //データ取得
                sda = new SqlDataAdapter(cmd);
                sda.Fill(comboDt3);

                //③窓口情報取得
                //SQL生成
                cmd.CommandText = "SELECT  " +
                " T2.MadoguchiShinchokuJoukyou " +               //0:進捗状況
                ",T2.MadoguchiID " +                             //1:窓口ID
                //",T2.MadoguchiUketsukeBangou " +                 //2:特調番号
                ",CASE WHEN T2.MadoguchiUketsukeBangouEdaban is null OR T2.MadoguchiUketsukeBangouEdaban = '' then T2.MadoguchiUketsukeBangou " + // 2:特調番号 + 枝番（存在すれば）
                " ELSE T2.MadoguchiUketsukeBangou + '-' + T2.MadoguchiUketsukeBangouEdaban END AS tokuchoNo " +
                ",T2.MadoguchiHachuuKikanmei " +                 //3:発注者詳細名
                ",T2.MadoguchiHikiwatsahi " +                    //4:遠隔地引渡承認
                ",T2.MadoguchiSaishuuKensa " +                   //5:遠隔地最終検査
                ",T2.MadoguchiJutakubi " +                       //6:受託日(依頼日)
                //",CASE WHEN T2.MadoguchiChousaKubunShibuHonbu= 1 THEN T1.OuenJoukyou ELSE 0 END " + //7:報告完了メールボタン → 応援状況
                ",CASE WHEN T2.MadoguchiChousaKubunShibuHonbu= 1 THEN CASE WHEN T1.OuenJoukyou = 2 THEN 2 ELSE 1 END ELSE 0 END " + //7:報告完了メールボタン → 応援状況
                ",CASE WHEN T2.MadoguchiChousaKubunShibuHonbu= 1 THEN T1.OuenKanryou ELSE 0 END " + //8:応援依頼             → 応援完了
                ",T2.MadoguchiTourokubi " +                      //9:登録日
                ",T2.MadoguchiOuenUketsukebi " +                 //10:応援受付日
                //",T1.OuenUketsukeDate " +                        //10:応援受付日
                ",T2.MadoguchiShimekiribi " +                    //11:締切日
                ",T2.MadoguchiHoukokuJisshibi " +                //12:報告実施日
                //",T2.MadoguchiChousaShubetsu " +                 //13:調査種別
                //",T2.MadoguchiJiishiKubun " +                    //14:実施区分
                ",CASE T2.MadoguchiChousaShubetsu WHEN 1 THEN '単品' WHEN 2 THEN '一般' WHEN 3 THEN '単契'ELSE ' ' END AS MadoguchiChousaShubetsu" +                      //13:調査種別
                ",CASE T2.MadoguchiJiishiKubun WHEN 1 THEN '実施' WHEN 2 THEN '打診中' WHEN 3 THEN '中止' ELSE ' ' END AS MadoguchiJiishiKubun" +       //14:実施区分
                ",T2.MadoguchiJutakuBushoCD  " +                 //15:受託部所
                ",T5.ChousainMei AS MadoguchiJutakuTantousha " + //16:契約担当者
                //",T3.ShibuMei AS MadoguchiTantoushaBushomei " +  //17:窓口部所1
                ",T2.MadoguchiTantoushaBushoCD AS MadoguchiTantoushaBushoCD " + //17:窓口部所1
                ",T7.ChousainMei AS MadoguchiTantousha " +       //18:窓口担当者1
                //",T2.MadoguchiChousaKubun " +                    //調査区分
                //",T2.MadoguchiChousaKubunJibusho " +             //19:調査区分 自部所
                //",T2.MadoguchiChousaKubunShibuShibu " +          //20:調査区分 支→支
                //",T2.MadoguchiChousaKubunHonbuShibu " +          //21:調査区分 本→支
                //",T2.MadoguchiChousaKubunShibuHonbu " +          //22:調査区分 支→本
                ",CASE T2.MadoguchiChousaKubunJibusho WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunJibusho" +        //19:調査区分 自部所
                ",CASE T2.MadoguchiChousaKubunShibuShibu WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunShibuShibu" +  //20:調査区分 支→支
                ",CASE T2.MadoguchiChousaKubunHonbuShibu WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunHonbuShibu" +  //21:調査区分 本→支
                ",CASE T2.MadoguchiChousaKubunShibuHonbu WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunShibuHonbu" +  //22:調査区分 支→本
                //",T2.MadoguchiGyoumuKanrishaCD AS MadoguchiGyoumuKanrishaCD " + //23:業務管理者
                ",T9.ChousainMei AS MadoguchiGyoumuKanrisha " +  //23:業務管理者
                ",T2.MadoguchiKanriBangou " +                    //24:管理番号
                ",T2.MadoguchiGyoumuMeishou " +                  //25:業務名称
                ",T2.MadoguchiKoujiKenmei " +                    //26:工事件名
                ",T2.MadoguchiBikou " +                          //27:備考
                ",T2.MadoguchiChousaHinmoku ";                  //28:調査品目
                //",T2.MadoguchiBushoRenban ";                     //29:部所連番


                //",'' " +                                         //30:備考1
                //",'' " +                                         //31:備考2
                //",'' " +                                         //32:備考3
                //",'' " +                                         //33:備考4
                //",'' " +                                         //34:備考5
                //",'' " +                                         //35:備考6
                //",'' " +                                         //36:備考7
                //",'' " +                                         //37:備考8
                //",'' " +                                         //38:備考9
                //",'' " +                                         //39:備考10
                //",'' " +                                         //40:備考11
                //",'' " +                                         //41:備考12
                //",'' " +                                         //42:備考13
                //",'' " +                                         //43:備考14
                //",'' " +                                         //44:備考15
                //",'' " +                                         //45:備考16
                //",'' " +                                         //46:備考17
                //",'' " +                                         //47:備考18
                //",'' " +                                         //48:備考19
                //",'' " +                                         //49:備考20
                //",'' " +                                         //50:備考21
                //",'' " +                                         //51:備考22
                //",'' " +                                         //52:備考23
                //",'' " +                                         //53:備考24
                //",'' " +                                         //54:備考25
                DataRow shibudr;

                // 連番書き込みフラグ True:書き込み False:書いてない
                Boolean renbanFlg = false;
                // 29:部所連番
                for (int i = 0; i < 25; i++)
                {
                    // 支部備考のデータがあるかないか
                    if (comboDt3.Rows.Count > i)
                    {
                        if (item_MadoguchiBusho.SelectedValue != null && item_MadoguchiBusho.SelectedValue.ToString() == comboDt3.Rows[i][0].ToString())
                        {
                            renbanFlg = true;
                            cmd.CommandText += ",(select TOP 1 ShibuBikouKanriNo from ShibuBikou where ShibuBikouBushoKanriboBushoCD = " + item_MadoguchiBusho.SelectedValue + " and MadoguchiID = T1.MadoguchiID) AS BushoRenban ";
                            break;
                        }
                    }
                }
                if (renbanFlg == false)
                {
                    cmd.CommandText += ",'' AS BushoRenban ";
                }

                // 備考1～備考25
                for (int i = 0; i < 25; i++)
                {
                    // 支部備考のデータがあるかないか
                    if (comboDt3.Rows.Count > i)
                    {
                        shibudr = comboDt3.Rows[i];
                        cmd.CommandText += ",(select TOP 1 ShibuBikouChousaBusho from ShibuBikou where ShibuBikouBushoKanriboBushoCD = " + shibudr["GyoumuBushoCD"] + " and MadoguchiID = T1.MadoguchiID) ";
                    }
                    else
                    {
                        cmd.CommandText += ",'' ";
                    }
                }

                cmd.CommandText +=
                ",T2.MadoguchiHachuuTantousha " +                //55:発注担当者
                ",T2.MadoguchiHachuuMail " +                     //56:メール
                ",T2.MadoguchiHachuuTEL " +                      //57:電話
                ",T2.MadoguchiRank " +                           //58:ランク
                ",T2.MadoguchiSaishuuKensaCheck " +              //59:最終検査
                ",T2.MadoguchiTourokuNendo " +                   //60:登録年度
                //",T1.OuenUketsukeID " +
                //",T2.MadoguchiTantoushaBushoCD AS MadoguchiTantoushaBushoCD " +
                //",T2.MadoguchiJutakuBushoCD AS MadoguchiJutakuBushoCD " +
                //",T2.MadoguchiHonbuTanpinflg " +
                //",T1.OuenUketsukeDate " +
                //",T2.MadoguchiKanryou " +
                //",T2.MadoguchiShouninsha " +
                //",T2.MadoguchiShouninnbi " +
                //",T2.MadoguchiMitsumoriTeishutu " +
                //",T2.MadoguchiTeiNyuusatsu " +
                //",T4.ShibuMei AS MadoguchiJutakubushoMei " +
                //",T2.MadoguchiJutakuTantoushaID AS MadoguchiJutakuTantoushaID " +
                //",T2.JutakuBushoShozokuCD AS JutakuBushoShozokuCD " +
                //",T6.BushoShozokuChou AS JutakuBushoShozokuChou " +
                //",T2.MadoguchiBusho " +
                //",T3.ShibuMei AS MadoguchiTantoushaBushomei " +
                //",T2.MadoguchiTantoushaCD AS MadoguchiTantoushaCD " +
                //",T2.MadoguchiBushoShozokuCD AS MadoguchiBushoShozokuCD " +
                //",T8.BushoShozokuChou AS MadoguchiBushoShozokuchou " +
                //",T9.ChousainMei AS MadoguchiGyoumuKanrisha " +
                //",T2.MadoguchiJutakuBangou " +
                //",T2.MadoguchiJutakuBangouEdaban " +
                //",T2.MadoguchiAnkenJouhouID " +
                //",T2.MadoguchiHachuukikanCD AS MadoguchiHachuukikanCD " +
                //",T10.HachushaMei AS MadoguchiHachuushamei " +
                //",T2.MadoguchiTankaTekiyou " +
                //",T2.MadoguchiNiwatashi " +
                //",T2.MadoguchiGyoumuRenrakuhyou " +
                //",T2.MadoguchiShiryouHolder " +
                //",T2.MadoguchiHoukokuzumi " +
                //",T2.MadoguchiHachuubusho " +
                //",T2.MadoguchiOldBushoflg " +
                //",T2.MadoguchiUketsukeBangouEdaban " +
                // 61:進捗アイコンの判定用
                ", " +
                "CASE " +
                "WHEN T2.MadoguchiHoukokuzumi = 1 THEN '8' " +
                "WHEN T2.MadoguchiHoukokuzumi != 1 THEN " +
                "     CASE " +
                "         WHEN T2.MadoguchiShinchokuJoukyou = 80 THEN '6' " +
                "         WHEN T2.MadoguchiShinchokuJoukyou = 70 THEN '5' " +
                "         WHEN T2.MadoguchiShinchokuJoukyou = 50 THEN '7' " +
                "         WHEN T2.MadoguchiShinchokuJoukyou = 60 THEN '7' " +
                "     ELSE " +
                "         CASE " +
                "              WHEN T2.MadoguchiShimekiribi < '" + DateTime.Today + "' THEN '1' " +
                "              WHEN T2.MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(3) + "' THEN '2' " +
                "              WHEN T2.MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(7) + "' THEN '3' " +
                "         ELSE '4' " +
                "         END " +
                "     END " +
                "END " +
                //"FROM OuenUketsuke T1 " +
                //" LEFT JOIN MadoguchiJouhou T2 ON T2.MadoguchiID = T1.MadoguchiID " +
                // GeneXusではOuneUketsukeに対してJOINしていたが、応援受付は新規の時に入れていないので、MadoguchiJouhouをベースに繋げる
                "FROM MadoguchiJouhou T2 " +
                " LEFT JOIN OuenUketsuke T1 ON T2.MadoguchiID = T1.MadoguchiID " +
                " LEFT JOIN Mst_Busho T3 ON T3.GyoumuBushoCD = T2.MadoguchiTantoushaBushoCD  " +
                " LEFT JOIN Mst_Busho T4 ON T4.GyoumuBushoCD = T2.MadoguchiJutakuBushoCD  " +
                " LEFT JOIN Mst_Chousain T5 ON T5.KojinCD = T2.MadoguchiJutakuTantoushaID  " +
                " LEFT JOIN Mst_Busho T6 ON T6.GyoumuBushoCD = T2.JutakuBushoShozokuCD  " +
                " LEFT JOIN Mst_Chousain T7 ON T7.KojinCD = T2.MadoguchiTantoushaCD  " +
                " LEFT JOIN Mst_Busho T8 ON T8.GyoumuBushoCD = T2.MadoguchiBushoShozokuCD  " +
                " LEFT JOIN Mst_Chousain T9 ON T9.KojinCD = T2.MadoguchiGyoumuKanrishaCD  " +
                " LEFT JOIN Mst_Hachusha T10 ON T10.HachushaCD = T2.MadoguchiHachuukikanCD ";

                // 調査担当部所 が選択されている場合、MadoguchiJouhouMadoguchiL1ChouをJOINする
                if (item_ChousaBusho.Text != "" && item_ChousaBusho.SelectedValue != null) {
                    cmd.CommandText +=
                   //" LEFT JOIN MadoguchiJouhouMadoguchiL1Chou T11 ON T2.MadoguchiID = T11.MadoguchiID ";

                   // 同じ部所で担当者が複数入れた場合、同一列が複数出てしまわないようにする為、以下のカタチとする
                   "  LEFT JOIN " +
                   //"(select top 1 MadoguchiID, MadoguchiL1ChousaBushoCD from MadoguchiJouhouMadoguchiL1Chou WHERE MadoguchiL1ChousaBushoCD = " + item_ChousaBusho.SelectedValue.ToString() + ") T11 " + 
                   //"    ON T2.MadoguchiID = T11.MadoguchiID ";

                   // 20210331修正 担当部所を部所CDで引っ掛けた際の1件目しか持ってきていなかったのを下記のカタチに修正
                    "(select MadoguchiID, MadoguchiL1ChousaBushoCD from MadoguchiJouhouMadoguchiL1Chou WHERE MadoguchiL1ChousaBushoCD = " + item_ChousaBusho.SelectedValue.ToString() + " GROUP BY MadoguchiID, MadoguchiL1ChousaBushoCD) T11 " +
                    "    ON T2.MadoguchiID = T11.MadoguchiID AND T11.MadoguchiL1ChousaBushoCD = '" + item_ChousaBusho.SelectedValue.ToString() + "' ";
                }
                // 指示書 が選択されている場合、TanpinNyuuryokuをJOINする
                if (item_Shijisho.Text != "")
                {
                    cmd.CommandText +=
                   " LEFT JOIN TanpinNyuuryoku T12 ON T2.MadoguchiID = T12.MadoguchiID ";
                }

                cmd.CommandText +=
               "WHERE MadoguchiTourokuNendo <= '" + nendo1 + "' and MadoguchiTourokuNendo >= '" + nendo2 + "' " +
                " AND MadoguchiDeleteFlag != 1 " +
                " AND T2.MadoguchiSystemRenban > 0 ";

                String w_jyokyou = "";
                String workdayFrom = "0";
                String workdayTo = "0";
                DateTime w_Simekiribi6 = DateTime.Today;

                // 締切日From,Toは後で計算値を設定
                DateTime w_SimekiribiFrom = DateTime.Today;
                DateTime w_SimekiribiTo = DateTime.Today;
                DateTime dateTime;
                DateTime w_Simekiribi7 = DateTime.MinValue;

                w_Simekiribi6 = w_Simekiribi6.AddDays(6);

                String Where = "";

                // 進捗状況
                if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() != "") 
                { 
                    switch (item_Shintyokujyoukyo.SelectedValue.ToString())
                    {
                        // 2次検証済み
                        // 完了
                        case "1":
                            //w_jyokyou = "3";
                            // 3:二次検済　⇒　70：二次検済
                            w_jyokyou = "70";
                            break;
                        // 担当者済
                        // 存在しない、ここは通らない
                        case "2":
                            //w_jyokyou = "2";
                            // 2:担当者済　⇒　50：担当者済
                            w_jyokyou = "50";
                            break;
                        // 締切日が1週間をこえる もしくは中止
                        // 締め切りまで1週間以上 または 中止
                        case "3":
                            w_Simekiribi7 = DateTime.Today.AddDays(7);
                            break;
                        // 締切日が1週間以内
                        // 締め切りまで1週間以内
                        case "4":
                            workdayFrom = "-1";
                            workdayTo = "7";
                            break;
                        // 締切日が3日以内
                        // 締め切りまで3日以内
                        case "5":
                            workdayFrom = "-1";
                            workdayTo = "3";
                            break;
                        // 締切日が超過
                        // 超過
                        case "6":
                            workdayTo = "-1";
                            break;
                        default:
                            break;
                    }
                }
                // 締切日付計算
                w_SimekiribiFrom = DateTime.Today.AddDays(int.Parse(workdayFrom));
                w_SimekiribiTo = DateTime.Today.AddDays(int.Parse(workdayTo));

                // 進捗状況のコンボボックスの条件　二次検証済み　か　完了かどうか
                if (w_jyokyou != "" && (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "1" || item_Shintyokujyoukyo.SelectedValue.ToString() == "2"))
                {
                    Where += "and MadoguchiShinchokuJoukyou = " + w_jyokyou + " ";
                }
                // //窓口完了が立っていないのが条件
                // 進捗状況が「締め切りまで1週間以上 または 中止（完了でない）」
                if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "3")
                {
                    //Where += "and (MadoguchiShimekiribi >= '" + w_Simekiribi6 + "') or(MadoguchiShinchokuJoukyou = 80) ";
                    //Where += "and MadoguchiShinchokuJoukyou < 50	or MadoguchiShinchokuJoukyou = 80 ";

                    //Where += "and ((MadoguchiShinchokuJoukyou < 50 and MadoguchiShimekiribi >= '" + w_Simekiribi6 + "') ";
                    //Where += "or MadoguchiShinchokuJoukyou = 80) ";
                    //Where += "and MadoguchiKanryou <> 1 ";
                    Where += "and ((MadoguchiShinchokuJoukyou < 50 and MadoguchiShimekiribi >= '" + w_Simekiribi7 + "') ";
                    Where += "or MadoguchiShinchokuJoukyou = 80) ";
                    //Where += "and MadoguchiKanryou <> 1 ";

                }
                // 進捗状況のコンボボックスの条件 締切日が1週間以内　3日以内　超過のとき
                if (workdayFrom != "0")
                {
                    Where += "and MadoguchiShimekiribi > '" + w_SimekiribiFrom + "' ";
                }
                if (workdayTo != "0")
                {
                    Where += "and MadoguchiShimekiribi <= '" + w_SimekiribiTo + "' ";
                }
                // 窓口完了が立っていないのが条件
                // 2次検証済ではない　のと　完了 中止ではないのが条件
                if (item_Shintyokujyoukyo.SelectedValue != null && (item_Shintyokujyoukyo.SelectedValue.ToString() == "4" || item_Shintyokujyoukyo.SelectedValue.ToString() == "5" || item_Shintyokujyoukyo.SelectedValue.ToString() == "6"))
                {
                    Where += "and MadoguchiKanryou <> 1 ";
                    Where += "and MadoguchiShinchokuJoukyou < 50 ";
                }
                // 1206 超過の場合は、報告済みは除外する
                if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "6")
                {
                    Where += "and T2.MadoguchiHoukokuzumi = 0 ";
                }
                // 特調番号
                if (item_TokuchoBangou.Text != "")
                {
                    //Where += "and (MadoguchiUketsukeBangou + '-' + MadoguchiUketsukeBangouEdaban) LIKE '%" + GlobalMethod.ChangeSqlText(item_TokuchoBangou.Text, 1, 0) + "%' ESCAPE '\\' ";

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
                        //if (OrAddFlg)
                        //{
                        //}
                        cmd.CommandText += " MadoguchiChousaKubunShibuShibu = 1 ";
                        //OrAddFlg = true;
                    }
                    else
                    {
                        cmd.CommandText += " MadoguchiChousaKubunShibuShibu = 0 ";
                    }
                    cmd.CommandText += "AND ";
                    //調査区分　本部→支部
                    if (item_ChousaKbnHonbuShibu.Checked)
                    {
                        //if (OrAddFlg)
                        //{
                        //        cmd.CommandText += "OR ";
                        //}
                        cmd.CommandText += " MadoguchiChousaKubunHonbuShibu = 1 ";
                        //OrAddFlg = true;
                    }
                    else
                    {
                        cmd.CommandText += " MadoguchiChousaKubunHonbuShibu = 0 ";
                    }

                    cmd.CommandText += "AND ";
                    //調査区分　支部→本部
                    if (item_ChousaKbnShibuHonbu.Checked)
                    {
                        //if (OrAddFlg)
                        //{
                        //    cmd.CommandText += "OR ";
                        //}
                        cmd.CommandText += " MadoguchiChousaKubunShibuHonbu = 1 ";
                        //OrAddFlg = true;
                    }
                    else
                    {
                        cmd.CommandText += " MadoguchiChousaKubunShibuHonbu = 0 ";
                    }
                    cmd.CommandText += ")";
                }



                // 登録日,締切日,受託日,報告実施日の選択
                // 4:登録日 1:締切日 2:受託日 3:報告実施日 5:応援受付日
                // 1:締切日への検索
                if (item_DateKindCombo.SelectedValue != null && item_DateKindCombo.SelectedValue.ToString() == "1" && item_DateFrom.CustomFormat == "")
                {
                    Where += "and MadoguchiShimekiribi >= '" + item_DateFrom.Text + "' ";
                }
                if (item_DateKindCombo.SelectedValue != null && item_DateKindCombo.SelectedValue.ToString() == "1" && item_DateTo.CustomFormat == "")
                {
                    dateTime = DateTime.Parse(item_DateTo.Text);
                    dateTime = dateTime.AddDays(1);
                    Where += "and MadoguchiShimekiribi < '" + dateTime + "' ";
                }
                // 4:登録日への検索
                if (item_DateKindCombo.SelectedValue != null && item_DateKindCombo.SelectedValue.ToString() == "4" && item_DateFrom.CustomFormat == "")
                {
                    Where += "and MadoguchiTourokubi >= '" + item_DateFrom.Text + "' ";
                }
                if (item_DateKindCombo.SelectedValue != null && item_DateKindCombo.SelectedValue.ToString() == "4" && item_DateTo.CustomFormat == "")
                {
                    dateTime = DateTime.Parse(item_DateTo.Text);
                    dateTime = dateTime.AddDays(1);
                    Where += "and MadoguchiTourokubi < '" + dateTime + "' ";
                }
                // 3:報告日への検索
                if (item_DateKindCombo.SelectedValue != null && item_DateKindCombo.SelectedValue.ToString() == "3" && item_DateFrom.CustomFormat == "")
                {
                    Where += "and MadoguchiHoukokuJisshibi >= '" + item_DateFrom.Text + "' ";
                }
                if (item_DateKindCombo.SelectedValue != null && item_DateKindCombo.SelectedValue.ToString() == "3" && item_DateTo.CustomFormat == "")
                {
                    dateTime = DateTime.Parse(item_DateTo.Text);
                    dateTime = dateTime.AddDays(1);
                    Where += "and MadoguchiHoukokuJisshibi < '" + dateTime + "' ";
                }
                // 2:受託日への検索
                if (item_DateKindCombo.SelectedValue != null && item_DateKindCombo.SelectedValue.ToString() == "2" && item_DateFrom.CustomFormat == "")
                {
                    Where += "and MadoguchiJutakubi >= '" + item_DateFrom.Text + "' ";
                }
                if (item_DateKindCombo.SelectedValue != null && item_DateKindCombo.SelectedValue.ToString() == "2" && item_DateTo.CustomFormat == "")
                {
                    dateTime = DateTime.Parse(item_DateTo.Text);
                    dateTime = dateTime.AddDays(1);
                    Where += "and MadoguchiJutakubi < '" + dateTime + "' ";
                }
                // 5:応援受付日への検索
                if (item_DateKindCombo.SelectedValue != null && item_DateKindCombo.SelectedValue.ToString() == "5" && item_DateFrom.CustomFormat == "")
                {
                    Where += "and OuenUketsukeDate >= '" + item_DateFrom.Text + "' ";
                }
                if (item_DateKindCombo.SelectedValue != null && item_DateKindCombo.SelectedValue.ToString() == "5" && item_DateTo.CustomFormat == "")
                {
                    dateTime = DateTime.Parse(item_DateTo.Text);
                    dateTime = dateTime.AddDays(1);
                    Where += "and OuenUketsukeDate < '" + dateTime + "' ";
                }
                // 発注者名・課名
                if (item_HachushaKamei.Text != "")
                {
                    Where += "and T2.MadoguchiHachuuKikanmei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_HachushaKamei.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                // 本部単品
                if (item_HonbuTanpin.Checked)
                {
                    Where += "and MadoguchiHonbuTanpinflg = 1 ";
                }
                // 受託部所
                if (item_JutakuBusho.Text != "")
                {
                    // 後ろ0を削った検索
                    Where += "and MadoguchiJutakuBushoCD LIKE '" + item_JutakuBusho.SelectedValue.ToString().TrimEnd('0') + "%'";
                }
                // 契約担当者
                if (item_KeiyakuTantousha.Text != "")
                {
                    //Where += "and MadoguchiJutakuTantousha LIKE '%" + GlobalMethod.ChangeSqlText(item_KeiyakuTantousha.Text, 1, 0) + "%' ESCAPE '\\' ";
                    Where += "and T5.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_KeiyakuTantousha.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                // 調査種別
                if (item_ChousaKind.SelectedValue != null && item_ChousaKind.Text != " ")
                {
                    Where += "and MadoguchiChousaShubetsu = " + item_ChousaKind.SelectedValue.ToString() + " ";
                }

                // 実施区分
                if (item_JishiKbn.SelectedValue != null && item_JishiKbn.SelectedValue.ToString() != "0")
                {
                    Where += "and T2.MadoguchiJiishiKubun = " + item_JishiKbn.SelectedValue.ToString() + " ";
                }
                // 窓口部所
                if (item_MadoguchiBusho.SelectedValue != null && item_MadoguchiBusho.Text != "")
                {
                    //MessageBox.Show("index:" + item_MadoguchiBusho.SelectedIndex);
                    //MessageBox.Show("index:" + item_MadoguchiBusho.Text);
                    if (item_MadoguchiBusho.SelectedIndex != 0)
                    {
                        Where += "and MadoguchiTantoushaBushoCD = " + item_MadoguchiBusho.SelectedValue.ToString() + " ";
                    }
                }
                // 窓口担当者
                if (item_MadoguchiTantousha.Text != "")
                {
                    //Where += "and MadoguchiTantousha LIKE '%" + GlobalMethod.ChangeSqlText(item_MadoguchiTantousha.Text, 1, 0) + "%' ESCAPE '\\' ";
                    Where += "and T7.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_MadoguchiTantousha.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                // 指示書 1:無し 2:有り
                if (item_Shijisho.SelectedValue != null && item_Shijisho.SelectedValue.ToString() == "1")
                {
                    Where += "and (TanpinShijisho = 0 or TanpinShijisho is null)";
                }
                // 指示書 1:無し 2:有り
                if (item_Shijisho.SelectedValue != null && item_Shijisho.SelectedValue.ToString() == "2")
                {
                    Where += "and TanpinShijisho = 1 ";
                }
                // 調査担当部所
                if(item_ChousaBusho.Text != "" && item_ChousaBusho.SelectedValue != null)
                {
                    Where += "and T11.MadoguchiL1ChousaBushoCD = " + item_ChousaBusho.SelectedValue.ToString() + " ";
                }
                if(Where != "")
                {
                    cmd.CommandText += Where;
                }
                cmd.CommandText += "ORDER BY ";
                //    "CASE " +
                //    "WHEN T2.MadoguchiHoukokuzumi = 1 THEN '8' " +
                //    "WHEN T2.MadoguchiHoukokuzumi != 1 THEN " +
                //    "     CASE " +
                //    "         WHEN T2.MadoguchiShinchokuJoukyou = 80 THEN '6' " +
                //    "         WHEN T2.MadoguchiShinchokuJoukyou = 70 THEN '5' " +
                //    "     ELSE " +
                //    "         CASE " +
                //    "              WHEN T2.MadoguchiShimekiribi < GETDATE() THEN '1' " +
                //    "              WHEN T2.MadoguchiShimekiribi <= DATEADD(day,3,GETDATE()) THEN '2' " +
                //    "              WHEN T2.MadoguchiShimekiribi <= DATEADD(day,7,GETDATE()) THEN '3' " +
                //    "         ELSE '4' " +
                //    "         END " +
                //    "     END " +
                //    "END " +
                //    " ,T1.OuenUketsukeID";


                // 完了は下にチェック時
                if (item_Kanryouhasita.Checked)
                {
                    cmd.CommandText += "CASE " +
                        "WHEN T2.MadoguchiHoukokuzumi = 1 THEN '8' " +
                        //"WHEN T2.MadoguchiHoukokuzumi != 1 THEN " +
                        //"     CASE " +
                        //"         WHEN T2.MadoguchiShinchokuJoukyou = 80 THEN '0' " +
                        //"     ELSE " +
                        //"         '1' " +
                        //"     END " +
                        "WHEN T2.MadoguchiHoukokuzumi != 1 THEN '0'" +
                        "END ,";

                    //cmd.CommandText += " ,T2.MadoguchiShimekiribi,CASE WHEN T2.MadoguchiUketsukeBangouEdaban is null OR T2.MadoguchiUketsukeBangouEdaban = '' then T2.MadoguchiUketsukeBangou END ";
                    //cmd.CommandText += " ,T2.MadoguchiShimekiribi,T2.MadoguchiUketsukeBangou,T2.MadoguchiUketsukeBangouEdaban ";

                }
                //else
                //{
                //    //cmd.CommandText += " T2.MadoguchiShimekiribi,CASE WHEN T2.MadoguchiUketsukeBangouEdaban is null OR T2.MadoguchiUketsukeBangouEdaban = '' then T2.MadoguchiUketsukeBangou END ";
                //    //cmd.CommandText += " T2.MadoguchiShimekiribi,T2.MadoguchiUketsukeBangou,T2.MadoguchiUketsukeBangouEdaban ";
                //}
                cmd.CommandText += " T2.MadoguchiShimekiribi,T2.MadoguchiUketsukeBangou,T2.MadoguchiUketsukeBangouEdaban ";


                Console.WriteLine(cmd.CommandText);
                GlobalMethod.outputLogger("Search_Madoguchi", "開始", "GetMadoguchiJouhou", UserInfos[1]);
                //データ取得
                sda = new SqlDataAdapter(cmd);

                if (ikkatsuFlg != "1") { 
                    set_error("", 0);
                }
                ListData.Clear();
                sda.Fill(ListData);

                // ヘッダー部だけに設定
                c1FlexGrid1.Rows.Count = 1;
                //// 行追加を有効 これが原因で1行多く表示されていた
                //c1FlexGrid1.AllowAddNew = true;

                //行数決め 基本は50
                int rowsCount = 50;
                if (ListData.Rows.Count <= 50)
                {
                    rowsCount = ListData.Rows.Count;
                }
                // 0件の場合
                if (ListData.Rows.Count == 0)
                {
                    if (ikkatsuFlg != "1")
                    {
                        set_error("", 0);
                    }
                    // I20001:該当データは0件です。
                    set_error(GlobalMethod.GetMessage("I20001", ""));
                }

                Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / int.Parse(item_Hyoujikensuu.Text.Replace("件", "")))).ToString();
                Paging_now.Text = (1).ToString();
                set_data(int.Parse(Paging_now.Text));
                set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));


                // 取得したデータをComponentOneのGridに入れていく処理
                //for (int i = 0; rowsCount > i; i++)
                //{
                //    // 行を追加
                //    c1FlexGrid1.Rows.Add();

                //    DataRow dr = ListData.Rows[i];

                //    // データをセットする[行,列]
                //    // 進捗状況
                //    //c1FlexGrid1[i + 1, 1] = dr[0];
                //    c1FlexGrid1[i + 1, 1] = dr[61]; // 進捗状況と締切日から計算した値をセット
                //    // 編集アイコンを表示する為に、1をセット
                //    c1FlexGrid1[i + 1, 2] = "1";

                //    // クエリ：1:窓口ID ～ 2:特調番号 ～ 29:部所連番 ～ 59:最終検査 
                //    // Grid：3:窓口ID ～ 4:特調番号 ～ 31:部所連番 ～ 61:最終検査
                //    int j = 0;
                //    for (j = 3; j <= 61; j++)
                //    {
                //        // j + 2 している理由は、Gridだと、0列目と編集アイコンがある分ズレる為
                //        c1FlexGrid1[i + 1, j] = dr[j - 2];
                //    }

                //    // 備考
                //    if(i == 0) { 
                //        // クエリ：30:備考1～54:備考25
                //        // Grid：32:備考1～56:備考25
                //        for (int k = 0; comboDt3.Rows.Count > k; k++)
                //        {
                //            DataRow dr2 = comboDt3.Rows[k];
                //            // BushokanriboKameiRaku の頭2文字をGridのヘッダーにセットする
                //            c1FlexGrid1[0, k + 32] = dr2[1];
                //        }
                //    }
                //}
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
                conn.Close();
            }
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
                for (int i = 0; i < c1FlexGrid1.Cols.Count; i++)
                {
                    // 進捗状況
                    c1FlexGrid1[r + 1, 1] = ListData.Rows[startrow + r][61];
                    // 編集アイコンを表示する為に、1をセット
                    c1FlexGrid1[r + 1, 2] = "1";
                    // 
                    c1FlexGrid1[r + 1, i + 2] = ListData.Rows[startrow + r][i];
                    // 59でループを抜ける
                    if (i == 59)
                    {
                        break;    
                    }
                }
                // 備考
                if (r == 0)
                {
                    // クエリ：30:備考1～54:備考25
                    // Grid：32:備考1～56:備考25
                    for (int k = 0; comboDt3.Rows.Count > k; k++)
                    {
                        DataRow dr2 = comboDt3.Rows[k];
                        // BushokanriboKameiRaku の頭2文字をGridのヘッダーにセットする
                        c1FlexGrid1[0, k + 32] = dr2[1];
                    }
                }
            }
            c1FlexGrid1.AllowAddNew = false;
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
            item_Nendo.DataSource = combodt;
            item_Nendo.DisplayMember = "Discript";
            item_Nendo.ValueMember = "Value";

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
            item_ChousaBusho.DataSource = combodt;
            item_ChousaBusho.DisplayMember = "Discript";
            item_ChousaBusho.ValueMember = "Value";


            DataTable combodt3 = GlobalMethod.getData(discript, value, table, where);
            dr = combodt3.NewRow();
            combodt3.Rows.InsertAt(dr, 0);

            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            sl = GlobalMethod.Get_SortedList(combodt3);
            c1FlexGrid1.Cols[19].DataMap = sl; // 窓口部所


            // Gridの受託部所の為の設定
            discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
            value = "Mst_Busho.GyoumuBushoCD ";
            table = "Mst_Busho";
            where = "";
            combodt = GlobalMethod.getData(discript, value, table, where);
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            c1FlexGrid1.Cols[17].DataMap = sl; // 受託部所

            // 登録日,締切日,受託日,報告実施日の選択
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(4, "登録日");
            tmpdt.Rows.Add(1, "締切日");
            tmpdt.Rows.Add(2, "受託日");
            tmpdt.Rows.Add(3, "報告実施日");
            tmpdt.Rows.Add(5, "応援受付日");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            item_DateKindCombo.DataSource = tmpdt;
            item_DateKindCombo.DisplayMember = "Discript";
            item_DateKindCombo.ValueMember = "Value";

            //受託部所
            discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
            value = "Mst_Busho.GyoumuBushoCD ";
            table = "Mst_Busho";
            where = "";
            combodt = GlobalMethod.getData(discript, value, table, where);
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            // c1FlexGrid1.Cols[20].DataMap = sl;


            // 窓口部所
            
            // 業務管理者
            discript = "Mst_Chousain.ChousainMei ";
            value = "Mst_Chousain.KojinCD ";
            table = "Mst_Chousain";
            where = "Mst_Chousain.ChousainDeleteFlag != 1 ";
            DataTable combodt4 = GlobalMethod.getData(discript, value, table, where);
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            sl = GlobalMethod.Get_SortedList(combodt4);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[25].DataMap = sl;



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
            item_FromTo.DataSource = tmpdt;
            item_FromTo.DisplayMember = "Discript";
            item_FromTo.ValueMember = "Value";

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
            item_ShimekiriSentaku.DataSource = tmpdt;
            item_ShimekiriSentaku.DisplayMember = "Discript";
            item_ShimekiriSentaku.ValueMember = "Value";

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
            item_ChousaKind.DataSource = tmpdt;
            item_ChousaKind.DisplayMember = "Discript";
            item_ChousaKind.ValueMember = "Value";
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[15].DataMap = sl; // 調査種別


            // 指示書
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(2, "有");
            tmpdt.Rows.Add(1, "無");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            item_Shijisho.DataSource = tmpdt;
            item_Shijisho.DisplayMember = "Discript";
            item_Shijisho.ValueMember = "Value";


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
            //if (tmpdt != null)
            //{
            //    dr = tmpdt.NewRow();
            //    tmpdt.Rows.InsertAt(dr, 0);
            //}
            item_JishiKbn.DataSource = tmpdt;
            item_JishiKbn.DisplayMember = "Discript";
            item_JishiKbn.ValueMember = "Value";
            //c1FlexGrid1.Cols[16].DataMap = sl; // 実施区分

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
            item_Shintyokujyoukyo.DataSource = tmpdt;
            item_Shintyokujyoukyo.DisplayMember = "Discript";
            item_Shintyokujyoukyo.ValueMember = "Value";

            // 管理帳票印刷



            // メール
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "未");
            tmpdt.Rows.Add(1, "済");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            c1FlexGrid1.Cols[54].DataMap = sl; // メール

            c1FlexGrid1.Cols[9].DataMap = sl; // 応援依頼
            c1FlexGrid1.Cols[10].DataMap = sl; // 応援完了

            // 調査区分
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "無");
            tmpdt.Rows.Add(1, "有");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            //c1FlexGrid1.Cols[21].DataMap = sl; // 調査区分 自部所
            //c1FlexGrid1.Cols[22].DataMap = sl; // 調査区分 支→支
            //c1FlexGrid1.Cols[23].DataMap = sl; // 調査区分 本→支
            //c1FlexGrid1.Cols[24].DataMap = sl; // 調査区分 支→本

            // 管理帳票印刷コンボボックス
            discript = "PrintName";
            value = "PrintListID";
            table = "Mst_PrintList";
            //where = "";
            where = "MENU_ID = 200 AND PrintBunruiCD = 1 AND PrintDelFlg <> 1 ORDER BY PrintListNarabijun";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            //dr = combodt.NewRow();
            //combodt.Rows.InsertAt(dr, 0);
            item_KanrityouhyouInsatu.DataSource = combodt;
            item_KanrityouhyouInsatu.DisplayMember = "Discript";
            item_KanrityouhyouInsatu.ValueMember = "Value";

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
            string JutakuBusho_SelectedValue = "";
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
            // 受託部所
            if (item_JutakuBusho.Text != "")
            {
                JutakuBusho_SelectedValue = item_JutakuBusho.SelectedValue.ToString();
            }

            // 調査担当部所
            string discript = "Mst_Busho.BushokanriboKamei ";
            string value = "Mst_Busho.GyoumuBushoCD ";
            string table = "Mst_Busho";
            string where = "KashoShibuCD != '' AND GyoumuBushoCD != '999990' AND GyoumuBushoCD != '127900' AND BushoNewOld <= 1 AND BushoMadoguchiHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 ";
            int FromNendo;
            int ToNendo;

            if (int.TryParse(nendo, out FromNendo))
            {
                ToNendo = int.Parse(nendo) + 1;
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
                item_ChousaBusho.DataSource = combodt;
                item_ChousaBusho.DisplayMember = "Discript";
                item_ChousaBusho.ValueMember = "Value";

                combodt2 = GlobalMethod.getData(discript, value, table, where);

                sl = new SortedList();
                //行の数だけの数だけSortedListにIDとValueをadd
                sl = GlobalMethod.Get_SortedList(combodt2);
                //該当グリッドのセルにセット
                //c1FlexGrid1.Cols[20].DataMap = sl;
                if (combodt2 != null)
                {
                    dr = combodt2.NewRow();
                    combodt2.Rows.InsertAt(dr, 0);
                }

                // 窓口部所
                item_MadoguchiBusho.DataSource = combodt2;
                item_MadoguchiBusho.DisplayMember = "Discript";
                item_MadoguchiBusho.ValueMember = "Value";

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
                item_JutakuBusho.DataSource = combodt;
                item_JutakuBusho.DisplayMember = "Discript";
                item_JutakuBusho.ValueMember = "Value";

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
                // 受託部所
                if (JutakuBusho_SelectedValue != "")
                {
                    item_JutakuBusho.SelectedValue = JutakuBusho_SelectedValue;
                }
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

            //VIPS　20220121　課題管理表No898　ADD　窓口ミハルの表示件数1000件対応
            //表示件数プルダウンの初期値
            int hyoujikensuuIndex = 3;
            int hyoujikensuuValue = 200;
            //マスタから取得
            String discript = "CommonMasterID";
            String value = "CommonValue1";
            String table = "M_CommonMaster";
            String where = "CommonMasterKye = 'MADOGUCHI_HINMOKU_HYOUJI_KENSUU'";
            DataTable dt = GlobalMethod.getData(discript, value, table, where);
            if (dt != null)
            {
                //1だったら2000件表示
                if (dt.Rows[0][0].ToString().Equals("1")) {
                    hyoujikensuuIndex = 4;
                    hyoujikensuuValue = 1000;
                }
            }

            item_Nendo.SelectedValue = GlobalMethod.GetTodayNendo();
            set_combo_shibu(item_Nendo.Text.ToString());

            //VIPS　20220121　課題管理表No898　CHANGE　窓口ミハルの表示件数1000件対応
            item_Hyoujikensuu.SelectedIndex = hyoujikensuuIndex;

            item_NendoOptionTounen.Checked = true;
            item_NendoOption3Nen.Checked = false;
            item_ChousaBusho.SelectedIndex = -1;
            item_HonbuTanpin.Checked = false;
            item_HachushaKamei.Text = "";
            item_TokuchoBangou.Text = "";
            item_DateKindCombo.SelectedIndex = -1;

            // 日付From
            item_DateFrom.Text = "";
            item_DateFrom.CustomFormat = " ";
            label_DateFrom.BackColor = Color.FromArgb(95, 158, 160);
            // 日付To
            item_DateTo.Text = "";
            item_DateTo.CustomFormat = " ";
            label_DateTo.BackColor = Color.FromArgb(95, 158, 160);

            item_JutakuBusho.SelectedIndex = -1;
            item_MadoguchiBusho.SelectedIndex = -1;
            item_ChousaKbnJibusho.Checked = false;
            item_ChousaKbnShibuShibu.Checked = false;
            item_ChousaKbnHonbuShibu.Checked = false;
            item_ChousaKbnShibuHonbu.Checked = false;
            item_Gyoumumei.Text = "";
            item_KanriBangou.Text = "";
            // 検索期間の選択
            item_FromTo.SelectedIndex = -1;
            item_KeiyakuTantousha.Text = "";
            item_MadoguchiTantousha.Text = "";
            item_Koujikenmei.Text = "";
            item_DashinGyoumu.Checked = false;
            item_ShimekiriSentaku.SelectedIndex = -1;
            //VIPS　20220121　課題管理表No898　CHANGE　窓口ミハルの表示件数1000件対応
            item_Hyoujikensuu.SelectedValue = hyoujikensuuValue;
            item_ChousaKind.SelectedIndex = -1;
            item_Shijisho.SelectedIndex = -1;
            item_JishiKbn.SelectedIndex = -1;
            item_ChousaHinmoku.Text = "";
            item_Kanryouhasita.Checked = true;
            item_Shintyokujyoukyo.SelectedIndex = -1;
            //item_KanrityouhyouInsatu.SelectedIndex = -1;


            //グリッドコントロールを初期化
            c1FlexGrid1.Styles.Normal.WordWrap = true;
            c1FlexGrid1.Rows[0].AllowMerging = true;
            c1FlexGrid1.AllowAddNew = false;

            if (c1FlexGrid1.Rows.Count > 1)
            {
                //グリッドクリア ヘッダー以外削除
                c1FlexGrid1.Rows.Count = 1;
            }
        }

        // 検索解除ボタン
        private void BtnClear_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            errorCheck_initialize();

            ClearForm();

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        // 検索ボタン
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            // false：正常 true：エラー
            Boolean errorFlg = false;

            //set_error("", 0);
            //// 日付の背景色を元に戻す
            //label_DateFrom.BackColor = Color.FromArgb(95, 158, 160);
            //label_DateTo.BackColor = Color.FromArgb(95, 158, 160);

            // 登録日、締切日～が選択されている
            //if (item_DateKindCombo.Text != null && item_DateKindCombo.Text != "")
            //{
            //    if (item_DateFrom.CustomFormat == "" && item_DateTo.CustomFormat == "")
            //    {
            //        // FromがToが大きい場合、エラー
            //        if (item_DateFrom.Value > item_DateTo.Value)
            //        {
            //            errorFlg = true;
            //            // E20002 対象項目の入力に誤りがあります。
            //            set_error(GlobalMethod.GetMessage("E20002", ""));
            //            //item_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //            //item_DateTo.BackColor = Color.FromArgb(255, 204, 255);
            //            // ピンクに設定
            //            label_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //            label_DateTo.BackColor = Color.FromArgb(255, 204, 255);
            //        }
            //        else
            //        {
            //            //item_DateFrom.BackColor = Color.FromArgb(255, 255, 255);
            //            //item_DateTo.BackColor = Color.FromArgb(255, 255, 255);
            //            label_DateFrom.BackColor = Color.FromArgb(95, 158, 160);
            //            label_DateTo.BackColor = Color.FromArgb(95, 158, 160);
            //        }
            //    }
            //}
            //// 登録日、締切日～が選択されていない
            //else
            //{
            //    if (item_DateFrom.CustomFormat == "" || item_DateTo.CustomFormat == "")
            //    {
            //        errorFlg = true;
            //        item_DateKindCombo.BackColor = Color.FromArgb(255, 204, 255);
            //        // E10010 必須入力項目が未入力です。赤背景の項目を入力して下さい。
            //        set_error(GlobalMethod.GetMessage("E10010", ""));
            //    }
            //}
            //// 検索期間の指定が空の場合
            //if(item_FromTo.Text != "")
            //{
            //    item_DateKindCombo.BackColor = Color.FromArgb(255, 255, 255);
            //    label_DateFrom.BackColor = Color.FromArgb(95, 158, 160);
            //}
            //else
            //{
            //    if (item_DateKindCombo.Text == "")
            //    {
            //        item_DateKindCombo.BackColor = Color.FromArgb(255, 204, 255);
            //    }
            //    else
            //    {
            //        item_DateKindCombo.BackColor = Color.FromArgb(255, 255, 255);
            //    }
            //    // 日付自が未選択
            //    if (item_DateFrom.CustomFormat != "")
            //    {
            //        label_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //    }
            //    else
            //    {
            //        label_DateFrom.BackColor = Color.FromArgb(95, 158, 160);
            //    }
            //}

            errorCheck_initialize();

            // 締切日の入力チェック
            if (item_DateFrom.CustomFormat == " " && item_DateTo.CustomFormat == " ")
            {
                // 検索期間の指定が空でない場合、エラー
                if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() != "")
                {
                    // 日付の種類が指定されていなかったら背景色を変える
                    if (item_DateKindCombo.Text == null || item_DateKindCombo.Text == "")
                    {
                        item_DateKindCombo.BackColor = errorBackColor;
                    }

                    errorFlg = true;
                    // E10010 必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", ""));
                    label_DateFrom.BackColor = errorBackColor;
                }
            }
            else
            {
                // 値の再設定
                if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() != "")
                {
                    changeShimekiribi();
                }

                // 日付の種類が指定されていなかったらエラー
                if (item_DateKindCombo.Text == null || item_DateKindCombo.Text == "")
                {
                    errorFlg = true;
                    // E10010 必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", ""));
                    item_DateKindCombo.BackColor = errorBackColor;
                }

                // 締切日が両方入力されていた場合
                if (item_DateFrom.CustomFormat == "" && item_DateTo.CustomFormat == "")
                {
                    // 締切日のFromとToの大小関係のチェック
                    if (item_DateFrom.Value > item_DateTo.Value)
                    {
                        errorFlg = true;
                        // E20002 対象項目の入力に誤りがあります。
                        set_error(GlobalMethod.GetMessage("E20002", ""));
                        label_DateFrom.BackColor = errorBackColor;
                        label_DateTo.BackColor = errorBackColor;
                    }
                }
            }

            if (errorFlg == false) { 
                get_data();
            }


            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        // 帳票出力ボタン
        private void BtnReport_Click(object sender, EventArgs e)
        {
            errorCheck_initialize();

            if (item_KanrityouhyouInsatu.Text == "")
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
                                    + " WHERE PrintListID = '" + item_KanrityouhyouInsatu.SelectedValue + "'"
                                    ;
                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(Dt);
                    Boolean errorFLG = false;
                    if (Dt.Rows.Count > 0)
                    {
                        set_error("", 0);
                        // 76:窓口ミハル一覧
                        if (Dt.Rows[0][0].ToString() == "76" || Dt.Rows[0][0].ToString() == "9")
                        {

                            if (errorFLG == false)
                            {
                                // string[]
                                // 29個分先に用意
                                string[] report_data = new string[29] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };

                                // 0.登録年度
                                report_data[0] = item_Nendo.SelectedValue.ToString();
                                // 1.登録年度オプション
                                if (item_NendoOptionTounen.Checked)
                                {
                                    report_data[1] = "1";   // 当年度
                                }
                                else
                                {
                                    report_data[1] = "2";   // 3年以内
                                }
                                // 2.調査担当部所CD
                                if (item_ChousaBusho.Text != null && item_ChousaBusho.Text != "")
                                {
                                    report_data[2] = item_ChousaBusho.SelectedValue.ToString();
                                }
                                // 3.本部単品
                                report_data[3] = "0";
                                if (item_HonbuTanpin.Checked)
                                {
                                    report_data[3] = "1";
                                }
                                // 4.発注者名・課名
                                report_data[4] = item_HachushaKamei.Text;
                                // 5.特調番号
                                report_data[5] = item_TokuchoBangou.Text;
                                // 6.日付選択
                                report_data[6] = "0";
                                if (item_DateKindCombo.Text != null && item_DateKindCombo.Text != "")
                                {
                                    report_data[6] = item_DateKindCombo.SelectedValue.ToString();
                                }
                                // 7.締切日from
                                report_data[7] = "null";
                                if (item_DateFrom.CustomFormat == "")
                                {
                                    report_data[7] = "'" + item_DateFrom.Text + "'";
                                }
                                // 8.締切日to
                                report_data[8] = "null";
                                if (item_DateTo.CustomFormat == "")
                                {
                                    report_data[8] = "'" + item_DateTo.Text + "'";
                                }
                                // 9.受託部所CD
                                if (item_JutakuBusho.Text != null && item_JutakuBusho.Text != "")
                                {
                                    report_data[9] = item_JutakuBusho.SelectedValue.ToString();
                                }
                                // 10.契約担当者
                                report_data[10] = item_KeiyakuTantousha.Text;
                                // 11.窓口部所CD
                                if (item_MadoguchiBusho.Text != null && item_MadoguchiBusho.Text != "")
                                {
                                    report_data[11] = item_MadoguchiBusho.SelectedValue.ToString();
                                }
                                // 12.窓口担当者
                                report_data[12] = item_MadoguchiTantousha.Text;
                                // 13.調査区分（自部所）
                                report_data[13] = "0";
                                if (item_ChousaKbnJibusho.Checked)
                                {
                                    report_data[13] = "1";
                                }
                                // 14.調査区分（支→支）
                                report_data[14] = "0";
                                if (item_ChousaKbnShibuShibu.Checked)
                                {
                                    report_data[14] = "1";
                                }
                                // 15.調査区分（本→支）
                                report_data[15] = "0";
                                if (item_ChousaKbnHonbuShibu.Checked)
                                {
                                    report_data[15] = "1";
                                }
                                // 16.調査区分（支→本）
                                report_data[16] = "0";
                                if (item_ChousaKbnShibuHonbu.Checked)
                                {
                                    report_data[16] = "1";
                                }
                                // 17.業務名称
                                report_data[17] = item_Gyoumumei.Text;
                                // 18.管理番号
                                report_data[18] = item_KanriBangou.Text;
                                // 19.工事件名
                                report_data[19] = item_Koujikenmei.Text;
                                // 20.打診中の業務
                                report_data[20] = "0";
                                if (item_ChousaKbnShibuHonbu.Checked)
                                {
                                    report_data[20] = "1";
                                }
                                // 21.検索期間の指定
                                report_data[21] = "0";
                                if (item_FromTo.Text != null && item_FromTo.Text != "")
                                {
                                    report_data[21] = item_FromTo.SelectedValue.ToString();
                                }
                                // 22.締め日の選択
                                report_data[22] = "0";
                                if (item_ShimekiriSentaku.Text != null && item_ShimekiriSentaku.Text != "")
                                {
                                    report_data[22] = item_ShimekiriSentaku.SelectedValue.ToString();
                                }
                                // 23.調査種別
                                report_data[23] = "0";
                                if (item_ChousaKind.Text != null && item_ChousaKind.Text != "")
                                {
                                    report_data[23] = item_ChousaKind.SelectedValue.ToString();
                                }
                                // 24.指示書
                                report_data[24] = "0";
                                if (item_Shijisho.Text != null && item_Shijisho.Text != "")
                                {
                                    report_data[24] = item_Shijisho.SelectedValue.ToString();
                                }
                                // 25.実施区分
                                report_data[25] = "0";
                                //不具合No1364（1145）
                                //実施区分のコンボじゃなく、指示書を参照していたため修正
                                if (item_JishiKbn.Text != null && item_JishiKbn.Text != "")
                                {
                                    report_data[25] = item_JishiKbn.SelectedValue.ToString();
                                }
                                //if (item_Shijisho.Text != null && item_Shijisho.Text != "")
                                //{
                                //    report_data[25] = item_Shijisho.SelectedValue.ToString();
                                //}
                                // 26.調査品目
                                report_data[26] = item_ChousaHinmoku.Text;
                                // 27.完了は下
                                report_data[27] = "0";
                                if (item_Kanryouhasita.Checked)
                                {
                                    report_data[27] = "1";
                                }
                                // 28.進捗状況
                                report_data[28] = "0";
                                if (item_Shintyokujyoukyo.Text != null && item_Shintyokujyoukyo.Text != "")
                                {
                                    report_data[28] = item_Shintyokujyoukyo.SelectedValue.ToString();
                                }

                                int PrintListID = 0;
                                string printName = "";
                                if (Dt.Rows[0][0].ToString() == "76")
                                {
                                    PrintListID = 231;
                                    printName = "MiharuIchiran";
                                }
                                else if (Dt.Rows[0][0].ToString() == "9")
                                {
                                    PrintListID = 12;
                                    printName = "KanriChouhyou";
                                }
                                string[] result = GlobalMethod.InsertMadoguchiReportWork(PrintListID, UserInfos[0], report_data, printName);


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
                                    // E00091:エラーが発生しました
                                    set_error(GlobalMethod.GetMessage("E00091", ""));
                                }
                            }
                        }

                        // 80:窓口ミハル一括取込用
                        if (Dt.Rows[0][0].ToString() == "80")
                        {
                            // string[]
                            // 29個分先に用意
                            string[] report_data = new string[1] { "" };

                            // 0.ダミーデータ
                            report_data[0] = "0";

                            string[] result = GlobalMethod.InsertMadoguchiReportWork(711, UserInfos[0], report_data, "IkkatsuTorikomi");

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
                                // E00091:エラーが発生しました
                                set_error(GlobalMethod.GetMessage("E00091", ""));
                            }

                        }
                    }
                    conn.Close();
                }
            }
        }

        // 登録年度変更
        private void item_Nendo_TextChanged(object sender, EventArgs e)
        {
            set_combo_shibu(item_Nendo.SelectedValue.ToString());
        }
        // 登録年度オプションクリック
        private void item_NendoOption_Click(object sender, EventArgs e)
        {
            set_combo_shibu(item_Nendo.SelectedValue.ToString());
        }

        // Gridの編集セルをクリックした場合の動作
        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            // マウスがクリックしたセル
            var hti = this.c1FlexGrid1.HitTest(new System.Drawing.Point(e.X, e.Y));
            var _row = hti.Row;
            var _col = hti.Column;

            // 編集を押されている、
            //if (hti.Column == 2 && hti.Row > 1 && c1FlexGrid1[hti.Row, hti.Column].ToString() == "1")
            if (hti.Column == 2 && hti.Row >= 1)
            {
                string mode;
                mode = "update";

                //// Role:2システム管理者
                //if (UserInfos[4].Equals("2"))
                //{
                //    mode = "";
                //}
                //if (UserInfos[2].Equals(c1FlexGrid1[hti.Row, 20].ToString()))
                //{
                //    mode = "";
                //}
                //else
                //{

                //}
                this.ReSearch = true;
                string MadoguchiID = c1FlexGrid1[hti.Row, hti.Column + 1].ToString();

                GlobalMethod.outputLogger("Madoguchi", "Madoguchi_Input インスタンス生成開始:" + DateTime.Now.ToString(), "", "DEBUG");

                Madoguchi_Input form = new Madoguchi_Input();
                GlobalMethod.outputLogger("Madoguchi", "Madoguchi_Input インスタンス生成終了:" + DateTime.Now.ToString(), "", "DEBUG");

                form.mode = mode;
                form.MadoguchiID = MadoguchiID;
                form.UserInfos = this.UserInfos;

                GlobalMethod.outputLogger("Madoguchi", "Madoguchi_Input Form.Show開始:" + DateTime.Now.ToString(), "", "DEBUG");

                form.Show(this);
                GlobalMethod.outputLogger("Madoguchi", "Madoguchi_Input Form.Show終了:" + DateTime.Now.ToString(), "", "DEBUG");

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

        // 日付系の値変更（CustomFormatを空文字にすると表示される）
        private void dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            ((DateTimePicker)sender).CustomFormat = "";
        }

        // 日付系の値変更（CustomFormatを空文字にすると表示される）
        private void dateTimePicker_CloseUp(object sender, EventArgs e)
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

        // 契約担当者プロンプト
        private void KeiyakuTantouhsa_Prompt_Click(object sender, EventArgs e)
        {
            Popup_ChousainList form = new Popup_ChousainList();
            //選択されている年度を条件に調査員プロンプトを表示
            if (item_Nendo.Text != "")
            {
                form.nendo = item_Nendo.SelectedValue.ToString();
            }
            form.program = "madoguchi";
            //form.Busho = UserInfos[2];
            form.Busho = null;
            if (item_JutakuBusho.Text != "")
            {
                form.Busho = item_JutakuBusho.SelectedValue.ToString();
            }
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item_KeiyakuTantousha.Text = form.ReturnValue[1];
                item_JutakuBusho.SelectedValue = form.ReturnValue[2];
            }
        }
        // 窓口担当者プロンプト
        private void MadoguchiTantouhsa_Prompt_Click(object sender, EventArgs e)
        {
            Popup_ChousainList form = new Popup_ChousainList();
            //選択されている年度を条件に調査員プロンプトを表示
            if (item_Nendo.Text != "")
            {
                form.nendo = item_Nendo.SelectedValue.ToString();
            }
            form.program = "madoguchi";
            //form.Busho = UserInfos[2];
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
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            Paging_now.Text = (1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        // 一つ前ページへ
        private void Previous_Page_Click(object sender, EventArgs e)
        {            
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            Paging_now.Text = (int.Parse(Paging_now.Text) - 1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        // 次のページへ
        private void After_Page_Click(object sender, EventArgs e)
        {            
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            Paging_now.Text = (int.Parse(Paging_now.Text) + 1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        // ラストページへ
        private void End_Page_Click(object sender, EventArgs e)
        {            
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            Paging_now.Text = (int.Parse(Paging_all.Text)).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }
        // ページングアイコンのON/OFF切替
        private void set_page_enabled(int now, int last)
        {
            GlobalMethod.outputLogger("Paging_Entry", "ページ:" + now, "GridAll", UserInfos[1]);
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

        // Activeになったときの処理
        private void Madoguchi_Search_Activated(object sender, EventArgs e)
        {
            if (ReSearch)
            {
                get_data();
                ReSearch = false;
            }
        }
        // 検索期間の指定変更
        private void Kikansitei_TextChanged(object sender, EventArgs e)
        {
            //kensakukikanFlg = true;
            //// 日付のFrom
            ////item_DateFrom.BackColor = Color.FromArgb(255, 255, 255);
            //label_DateFrom.BackColor = Color.FromArgb(95, 158, 160);
            //// 登録日、締切日、受託日、報告実施日の選択
            //item_DateKindCombo.BackColor = Color.FromArgb(255, 255, 255);

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
            //            //item_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //            label_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //        }
            //        else
            //        {
            //            //item_DateFrom.BackColor = Color.FromArgb(255, 255, 255);
            //            label_DateFrom.BackColor = Color.FromArgb(255, 255, 255);
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
            //            //item_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //            label_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //        }
            //        else
            //        {
            //            //item_DateFrom.BackColor = Color.FromArgb(255, 255, 255);
            //            label_DateFrom.BackColor = Color.FromArgb(95, 158, 160);
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
            //            //item_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //            label_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //        }
            //        else
            //        {
            //            DateTime dateTime = item_DateTo.Value;
            //            // Toに6日を引いた日をFromにセット
            //            dateTime = dateTime.AddDays(-6);
            //            //item_DateFrom.BackColor = Color.FromArgb(255, 255, 255)
            //            label_DateFrom.BackColor = Color.FromArgb(95, 158, 160);
            //            item_DateFrom.Value = dateTime;
            //        }
            //    }
            //}
            //// 登録日、締切日～が未選択かどうか
            //if (item_DateKindCombo.Text == "" && init_flg == false)
            //{
            //    item_DateKindCombo.BackColor = Color.FromArgb(255, 204, 255);
            //}
            //else
            //{
            //    item_DateKindCombo.BackColor = Color.FromArgb(255, 255, 255);
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

            //DateTime dateTime;
            //// 未選択の場合
            //if (item_ShimekiriSentaku.Text == "" && init_flg == false && kensakukikanFlg == false)
            //{
            //    item_DateKindCombo.BackColor = Color.FromArgb(255, 255, 255);
            //    // 検索期間の指定を空に
            //    item_FromTo.SelectedValue = -1;
            //    item_FromTo.BackColor = Color.FromArgb(255, 255, 255);
            //    //item_DateFrom.BackColor = Color.FromArgb(255, 255, 255);
            //    label_DateFrom.BackColor = Color.FromArgb(95, 158, 160);
            //}

            //// 1:本日の締めは？
            //if (item_ShimekiriSentaku.SelectedValue != null && item_ShimekiriSentaku.SelectedValue.ToString() == "1" && kensakukikanFlg == false)
            //{
            //    // 1:締切日に設定
            //    item_DateKindCombo.SelectedValue = 1;
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
            //else if (item_ShimekiriSentaku.SelectedValue != null && item_ShimekiriSentaku.SelectedValue.ToString() == "2" && kensakukikanFlg == false)
            //{
            //    // 1:締切日に設定
            //    item_DateKindCombo.SelectedValue = 1;
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
            //else if (item_ShimekiriSentaku.SelectedValue != null && item_ShimekiriSentaku.SelectedValue.ToString() == "3" && kensakukikanFlg == false)
            //{
            //    // 1:締切日に設定
            //    item_DateKindCombo.SelectedValue = 1;
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
            //kensakukikanFlg = false;

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
                    // 検索対象の日付の設定（1:締切日）
                    item_DateKindCombo.SelectedValue = 1;
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

        private void errorCheck_initialize()
        {
            // エラーメッセージのクリア
            set_error("", 0);

            // 画面背景色の初期化
            //label23.BackColor = Color.CadetBlue;
            item_DateKindCombo.BackColor = Color.White;
            label_DateFrom.BackColor = Color.CadetBlue;
            label_DateTo.BackColor = Color.CadetBlue;
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

            // 日付の種類が未選択だった場合、背景色の変更
            if (item_DateKindCombo.Text == null || item_DateKindCombo.Text == "")
            {
                item_DateKindCombo.BackColor = errorBackColor;
            }

            // 締切日が未入力だった場合、背景色の変更
            if (item_DateFrom.CustomFormat == " " && item_DateTo.CustomFormat == " ")
            {
                label_DateFrom.BackColor = errorBackColor;
            }
        }

        // VIPS　20220228　課題管理表No1259(949)　DEL　「窓口ミハル一括登録」ボタン非表示  対応
        //// 窓口ミハル一覧から新規登録
        //private void BtnMadoguchiIkkatsu_Click(object sender, EventArgs e)
        //{
        //    // 例外発生フラグ 0:発生なし 1:発生
        //    exceptionFlg = 0;

        //    // 採番の値をクリア
        //    saibanMadoguchiID = "";

        //    tokuchoBangou = "";


        //    Popup_Loading Loading = new Popup_Loading();
        //    Loading.StartPosition = FormStartPosition.CenterScreen;
        //    Loading.Show();

        //    errorId = 0;
        //    readCount = 0;
        //    errorUser = UserInfos[0] + "_" + DateTime.Today.ToString("yyyyMMdd");
        //    set_error("", 0);

        //    //ファイル
        //    OpenFileDialog Dialog1 = new OpenFileDialog();
        //    Dialog1.InitialDirectory = @"C:";
        //    Dialog1.Title = "インポートするファイルを選択してください。";

        //    if (Dialog1.ShowDialog() == DialogResult.OK)
        //    {
        //        get_excel(Dialog1.FileName);
        //    }
        //    else
        //    {
        //        // E70039:ファイルが読み込まれていません。
        //        set_error(GlobalMethod.GetMessage("E70039", ""));
        //    }

        //    Dialog1.Dispose();
        //    Loading.Close();

        //    // バックグラウンドとなっているExcelプロセスをKILL
        //    excelProcessKill();

        //    // データを取り直す ikkatsuFlg:1 一括フラグ:1 を渡すと画面のエラーメッセージをクリアしない
        //    get_data("1");
        //}


        private void get_excel(string fileName)
        {
            writeHistory("窓口ミハル一括登録を開始します。");

            sheetFlg1 = 0; // 調査概要タブ存在フラグ 0:あり 1:なし
            sheetFlg2 = 0; // 担当部所タブ存在フラグ 0:あり 1:なし
            sheetFlg3 = 0; // 調査品目一覧タブ存在フラグ 0:あり 1:なし
            sheetFlg4 = 0; // 協力依頼書タブ存在フラグ 0:あり 1:なし
            sheetFlg5 = 0; // 応援受付状況タブ存在フラグ 0:あり 1:なし
            sheetFlg6 = 0; // 単品入力項目タブ存在フラグ 0:あり 1:なし
            tokuchoBangou = "";

            // エラーフラグ true:エラー false:正常
            Boolean errorFlg = false;
            String ikkatsuFolder = "";
            String backupFolder = "";

            ikkatsuFolder = GlobalMethod.GetCommonValue1("MADOGUCHI_IKKATSU_FOLDER");
            backupFolder = GlobalMethod.GetCommonValue1("MADOGUCHI_IKKATSU_BACKUP_FOLDER");

            // 取込対象フォルダ
            if (ikkatsuFolder == null)
            {
                // E70058:取込対象フォルダが存在しません。
                set_error(GlobalMethod.GetMessage("E70058", ""));
                errorFlg = true;
            }
            else
            {
                // フォルダチェック
                if (!Directory.Exists(ikkatsuFolder))
                {
                    // E70058:取込対象フォルダが存在しません。
                    set_error(GlobalMethod.GetMessage("E70058", ""));
                    errorFlg = true;
                }
            }
            // 退避フォルダ
            if (backupFolder == null)
            {
                // E70059:退避フォルダが存在しません。
                set_error(GlobalMethod.GetMessage("E70059", ""));
                errorFlg = true;
            }
            else
            {
                // フォルダチェック
                if (!Directory.Exists(backupFolder))
                {
                    // E70059:退避フォルダが存在しません。。
                    set_error(GlobalMethod.GetMessage("E70059", ""));
                    errorFlg = true;
                }
            }
            // 取込対象フォルダ、退避フォルダが見つからない場合、ファイルを開く前に処理を終了する
            if (errorFlg == true)
            {
                return;
            }

            try
            {
                // 処理前にExcelプロセスを一旦全部killする
                excelProcessKill();
                // プロセスKILL後、1秒待機
                Thread.Sleep(1000);

                // 明示的にガベコレを行う
                GC.Collect();
                // ガベコレ後、1秒待機
                Thread.Sleep(1000);

                // ここでCOMExceptionが起こる場合があるらしい
                Application ExcelApp = null;
                Workbook wb = null;
                Worksheet ws = new Worksheet();

                try
                {
                    //Excel取込
                    ExcelApp = new Application();
                    wb = getExcelFile(fileName, ExcelApp);

                    if (wb == null)
                    {
                        // E70039:ファイルが読み込まれていません。
                        set_error(GlobalMethod.GetMessage("E70039", ""));
                        return;
                    }
                    // WorkBook Open フラグ 0:閉じている 1:開いている
                    workBookOpenFlg = 1;

                    //dynamic xlSheet = null;

                    //// シート存在チェック
                    //int sheetIndex = 0;
                    //string sheetName = "調査概要";
                    //sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                    //if (sheetIndex <= 0)
                    //{
                    //    // E70057:シートが存在しません。
                    //    set_error(GlobalMethod.GetMessage("E70057", sheetName));
                    //    errorFlg = true;
                    //}

                    //sheetName = "担当部所";
                    //sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                    //if (sheetIndex <= 0)
                    //{
                    //    // E70057:シートが存在しません。
                    //    set_error(GlobalMethod.GetMessage("E70057", sheetName));
                    //    errorFlg = true;
                    //}

                    //sheetName = "調査品目一覧";
                    //sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                    //if (sheetIndex <= 0)
                    //{
                    //    // E70057:シートが存在しません。
                    //    set_error(GlobalMethod.GetMessage("E70057", sheetName));
                    //    errorFlg = true;
                    //}

                    //sheetName = "協力依頼書";
                    //sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                    //if (sheetIndex <= 0)
                    //{
                    //    // E70057:シートが存在しません。
                    //    set_error(GlobalMethod.GetMessage("E70057", sheetName));
                    //    errorFlg = true;
                    //}

                    //// 施行条件・・・テンプレートから削除
                    ////sheetName = "施行条件";
                    ////sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                    ////if (sheetIndex <= 0)
                    ////{
                    ////    // E70057:シートが存在しません。
                    ////    set_error(GlobalMethod.GetMessage("E70057", sheetName));
                    ////    errorFlg = true;
                    ////}

                    //sheetName = "応援受付状況";
                    //sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                    //if (sheetIndex <= 0)
                    //{
                    //    // E70057:シートが存在しません。
                    //    set_error(GlobalMethod.GetMessage("E70057", sheetName));
                    //    errorFlg = true;
                    //}

                    //sheetName = "単品入力項目";
                    //sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                    //if (sheetIndex <= 0)
                    //{
                    //    // E70057:シートが存在しません。
                    //    set_error(GlobalMethod.GetMessage("E70057", sheetName));
                    //    errorFlg = true;
                    //}

                    // エラーフラグがfalseの場合、登録・更新処理
                    if (errorFlg == false)
                    {
                        // 登録更新処理
                        if (registrationMadoguchi(fileName, wb, ws))
                        {
                            // VIPS　20220228　課題管理表No1259(949)　DEL　「窓口ミハル一括登録取込み結果」ボタン非表示  対応
                            //BtnTorikomiKekka.Enabled = false; // Enabledを切るとC#の使用でForeColorが変更できないので背景のみ変更する
                            //BtnTorikomiKekka.BackColor = Color.DarkGray;

                            // E70038:取込が完了しました。
                            set_error(GlobalMethod.GetMessage("E70038", ""));
                        }
                        else
                        {
                            // 例外発生フラグ 0:発生なし 1:発生
                            if (exceptionFlg == 0)
                            {
                                // VIPS　20220228　課題管理表No1259(949)　DEL　「窓口ミハル一括登録取込み結果」ボタン非表示  対応
                                // E70037:取込ファイルにエラーがありました。
                                set_error(GlobalMethod.GetMessage("E70037", ""));
                                //BtnTorikomiKekka.Enabled = true;
                                //BtnTorikomiKekka.BackColor = Color.FromArgb(42, 78, 122);
                            }
                            else
                            {
                                // E00091:エラーが発生しました。
                                set_error(GlobalMethod.GetMessage("E00091", ""));
                            }
                        }
                    }
                    else
                    {
                        // E70037:取込ファイルにエラーがありました。
                        set_error(GlobalMethod.GetMessage("E70037", ""));
                        // VIPS　20220228　課題管理表No1259(949)　DEL　「窓口ミハル一括登録取込み結果」ボタン非表示  対応
                        //BtnTorikomiKekka.Enabled = true;
                        //BtnTorikomiKekka.BackColor = Color.FromArgb(42, 78, 122);
                    }
                }
                finally
                {
                    try 
                    { 
                        //Excelのオブジェクトを開放し忘れているとプロセスが落ちないため注意
                        Marshal.ReleaseComObject(ws);
                        //wb.Close(false);
                        if(workBookOpenFlg == 1)
                        {
                            //writeHistory("workBook close");
                            wb.Close(false, Type.Missing, Type.Missing);
                        }
                        Marshal.ReleaseComObject(wb);
                        if (ExcelApp != null)
                        {
                            ExcelApp.Quit();
                        }
                        Marshal.ReleaseComObject(ExcelApp);
                        ws = null;
                        wb = null;
                        ExcelApp = null;
                        GC.Collect();
                    }
                    catch (Exception excele)
                    {
                        writeHistory("Excelプロセスの開放エラー：" + excele.Message);
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException come)
            {
                // ExcelDDL関連エラー
                // E00090:システム障害が発生しました。
                set_error(GlobalMethod.GetMessage("E00090","EXCEL"));
                writeHistory("System.Runtime.InteropServices.COMException:" + come.Message);
            }
            tokuchoBangou = "";
            //saibanMadoguchiID = "";
            writeHistory("窓口ミハル一括登録処理を終了します。");
        }

        // 指定されたワークシート名のインデックスを返すメソッド
        private int getSheetIndex(string sheetName, Excel.Sheets shs)
        {
            int i = 0;
            foreach (Excel.Worksheet sh in shs)
            {
                if (sheetName == sh.Name)
                {
                    return i + 1;
                }
                i += 1;
            }
            return 0;
        }

        private Boolean registrationMadoguchi(string filePath,Workbook wb, Worksheet ws)
        {
            // ▼各シートのデータ情報
            // 調査概要 B2～B24
            // 担当部所     A～Bの3行目～
            // 調査品目一覧 A～ATの3行目～
            // 協力依頼書 B1～B15
            // --施工条件     B1～B113・・・テンプレートから削除
            // 単品入力項目 B1～B16

            Boolean registrationResult = true;

            try
            {
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    //エラーメッセージ
                    conn.Open();
                    var cmd = conn.CreateCommand();
                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    Excel.Range w_rgnName = null;

                    // データ配列（登録・更新ロジックに投げる際の配列）
                    // 調査概要
                    string[,] chousagaiyouSQLData = new string[1, 60];

                    // 担当部所 Garoon追加宛先
                    //string[,] tantoubushoGaroonAtesakiSQLData = new string[c1FlexGrid5.Rows.Count - 1, 5];
                    string[,] tantoubushoGaroonAtesakiSQLData = null;

                    // 調査品目明細・・・GeneXus呼び出しの為、無し

                    // 協力依頼書
                    string[,] kyouryokuIraishoSQLData = new string[1, 24];
                    // 応援受付
                    string[,] ouenUketsukeSQLData = new string[1, 5];
                    // 単品入力項目
                    string[,] tanpinNyuuryokuSQLData = new string[1, 37];

                    // ファイルパスからファイル名を取り出す
                    string fileName = "";
                    fileName = Path.GetFileName(filePath);

                    string workMsg = "";
                    Boolean workFlg = false;

                    String discript = "";
                    String value = "";
                    String table = "";
                    String where = "";
                    DataTable workDT = null;
                    string resultMsg = "";

                    string sheetName = "";

                    // 応援受付フラグ true：応援受付あり false:なし
                    Boolean ouenUketsukeFlg = true;

                    string kyouryokuIraishoID = "";
                    string kyourokuIraisakiTantoshaCD = "";

                    SqlDataAdapter kyouryokuIraiSda = null;

                    string tanpinNyuuryokuID = "";

                    SqlDataAdapter tanpinSda = null;

                    int FromNendo = 0;
                    int ToNendo = 0;

                    // Excelエラーフラグ true:エラー false：正常
                    Boolean excelErrorFlg = false;

                    int len = 0;
                    int col = 0;

                    try
                    {
                        dynamic xlSheet = null;

                        // シート存在チェック
                        // 調査概要以外はエラーとしない
                        int sheetIndex = 0;
                        sheetName = "調査概要";
                        sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                        if (sheetIndex <= 0)
                        {
                            // E70057:シートが存在しません。
                            //set_error(GlobalMethod.GetMessage("E70057", sheetName));
                            excelErrorFlg = true;
                            sheetFlg1 = 1;
                            return false;
                        }

                        sheetName = "担当部所";
                        sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                        if (sheetIndex <= 0)
                        {
                            // E70057:シートが存在しません。
                            //set_error(GlobalMethod.GetMessage("E70057", sheetName));
                            //errorFlg = true;
                            sheetFlg2 = 1;
                        }

                        sheetName = "調査品目一覧";
                        sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                        if (sheetIndex <= 0)
                        {
                            // E70057:シートが存在しません。
                            //set_error(GlobalMethod.GetMessage("E70057", sheetName));
                            //errorFlg = true;
                            sheetFlg3 = 1;
                        }

                        sheetName = "協力依頼書";
                        sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                        if (sheetIndex <= 0)
                        {
                            // E70057:シートが存在しません。
                            //set_error(GlobalMethod.GetMessage("E70057", sheetName));
                            //errorFlg = true;
                            sheetFlg4 = 1;
                        }

                        sheetName = "応援受付状況";
                        sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                        if (sheetIndex <= 0)
                        {
                            // E70057:シートが存在しません。
                            //set_error(GlobalMethod.GetMessage("E70057", sheetName));
                            //errorFlg = true;
                            sheetFlg5 = 1;
                        }

                        sheetName = "単品入力項目";
                        sheetIndex = getSheetIndex(sheetName, wb.Sheets);
                        if (sheetIndex <= 0)
                        {
                            // E70057:シートが存在しません。
                            //set_error(GlobalMethod.GetMessage("E70057", sheetName));
                            //errorFlg = true;
                            sheetFlg6 = 1;
                        }


                        // ▼調査概要
                        sheetName = "調査概要";
                        Worksheet worksheet = wb.Sheets[sheetName];
                        worksheet.Select();

                        workMsg = "";
                        workFlg = false;

                        xlSheet = null;
                        xlSheet = wb.Sheets[sheetName];
                        ws = wb.Sheets[sheetName];
                        w_rgnName = xlSheet.UsedRange;

                        // 使用されているエクセルの行数
                        int getRowCount = ws.UsedRange.Rows.Count;

                        // Excel操作用の配列の作成
                        // 二次元配列の各次元の最小要素番号
                        int[] lower = {1, 1};
                        // 二次元配列の各次元の要素数 
                        int[] length = {24, 1};
                        object[,] InputObject = (object[,])Array.CreateInstance(typeof(object), length, lower);

                        // B2から、B24を取得する（データは B2 からだが、index がズレると分かりづらくなるため、B1から取る）
                        Excel.Range InputRange = worksheet.Range[worksheet.Cells[1, 2], worksheet.Cells[24, 2]];
                        InputObject = (object[,])InputRange.Value;

                        ErrorMsg = "";

                        // 【調査概要】
                        //   共通ロジック                               Excelデータ
                        //-----------------------------------------------------------------------------------
                        //   00:画面モード(insert)
                        //   01:採番した窓口ID
                        // o 02:登録年度                               (string)w_rgnName[7, 2].Text;
                        //   03:遠隔地引渡承認
                        //   04:遠隔地最終検査
                        //   05:遠隔地承認者
                        //   06:遠隔地承認日
                        // o 07:調査担当者への締切日                 * (string)w_rgnName[22, 2].Text; → (string)w_rgnName[23, 2].Text;
                        // o 08:登録日                               * (string)w_rgnName[20, 2].Text; → (string)w_rgnName[21, 2].Text;
                        // o 09:報告実施日                             (string)w_rgnName[24, 2].Text; → (string)w_rgnName[25, 2].Text;
                        // o 10:調査種別                               (string)w_rgnName[16, 2].Text; → (string)w_rgnName[17, 2].Text;
                        // o 11:実施区分                               (string)w_rgnName[18, 2].Text; → (string)w_rgnName[19, 2].Text;
                        //   12:MadoguchiShinchokuJoukyou
                        //   13:受託課所支部
                        //   14:契約担当者orNULL
                        //   15:受託部所所属長の部所CD
                        // o 16:窓口部所                             * (string)w_rgnName[5, 2].Text; 
                        // o 17:窓口担当者                             (string)w_rgnName[6, 2].Text;
                        //   18:窓口部所所属長の部所CD
                        // o 19:調査区分　自部所                       (string)w_rgnName[11, 2].Text;
                        // o 20:調査区分　支→支                       (string)w_rgnName[12, 2].Text;
                        // o 21:調査区分　本→支                       (string)w_rgnName[13, 2].Text;
                        // o 22:調査区分　支→本                       (string)w_rgnName[14, 2].Text;
                        // o 23:管理番号                               (string)w_rgnName[10, 2].Text;
                        // o 24:受託番号                               (string)w_rgnName[8, 2].Text;
                        // o 25:受託番号枝番                           
                        //   26:特調番号                             
                        //   27:特調番号枝番                         * (string)w_rgnName[9, 2].Text;
                        //   28:発注者名・課名                                                           (string)w_rgnName[15, 2].Text;
                        //   29:業務名称
                        // o 30:工事件名                               (string)w_rgnName[15, 2].Text; → (string)w_rgnName[16, 2].Text;
                        // o 31:調査品目                               (string)w_rgnName[17, 2].Text; → (string)w_rgnName[18, 2].Text;
                        // o 32:備考                                   (string)w_rgnName[19, 2].Text; → (string)w_rgnName[20, 2].Text;
                        // o 33:単価適用地域                           (string)w_rgnName[21, 2].Text; → (string)w_rgnName[22, 2].Text;
                        // o 34:荷渡場所                               (string)w_rgnName[23, 2].Text; → (string)w_rgnName[24, 2].Text;
                        // o 35:報告済                                 (string)w_rgnName[2, 2].Text;
                        //   36:管理技術者
                        // o 37:本部単品                               (string)w_rgnName[3, 2].Text;
                        //   38:集計表フォルダ
                        //   39:報告書フォルダ
                        //   40:調査資料フォルダ
                        //   41:業務管理者の業務管理者CD                Null
                        //   42:AnkenJouhou.AnkenJouhouID
                        // o 43:MadoguchiGaroonRenkei                  (string)w_rgnName[4, 2].Text;
                        //   44:AnkenJouhou.AnkenJouhouID
                        //   45:連番

                        DataTable ankenJutakuDT = new DataTable();

                        // 受託番号が存在する場合
                        if ((string)w_rgnName[8, 2].Text != "")
                        {
                            cmd.CommandText = "SELECT " +
                                    "aj.AnkenTantoushaCD " +
                                    ",aj.AnkenTantoushaMei " +
                                    ",aj.AnkenAnkenBangou " +
                                    ",gj.KanriGijutsushaCD " +
                                    ",gj.KanriGijutsushaNM " +
                                    ",aj.AnkenJouhouID " +
                                    ",aj.AnkenJutakuBangou " +
                                    ",aj.AnkenJutakuBangouEda " +
                                    ",aj.AnkenKianZumi " +
                                    ",aj.AnkenHachuushaKaMei " +
                                    ",aj.AnkenGyoumuMei " +
                                    ",aj.GyoumuKanrishaCD " +
                                    ",aj.AnkenKeiyakusho " +
                                    //",gjm.GyoumuJouhouMadoGyoumuBushoCD " +
                                    ",mc.GyoumuBushoCD " +

                                    ",gjm.GyoumuJouhouMadoKojinCD " +
                                    ",gjm.GyoumuJouhouMadoChousainMei " +
                                    ",aj.AnkenJutakubushoCD " +
                                    "FROM AnkenJouhou aj " +
                                    "LEFT JOIN GyoumuJouhou gj ON " +
                                    "aj.AnkenJouhouID = gj.AnkenJouhouID " +
                                    "LEFT JOIN GyoumuJouhouMadoguchi gjm ON " +
                                    "gj.GyoumuJouhouID = gjm.GyoumuJouhouID " +

                                    // 1225 窓口部所は、窓口担当者のコードから調査員マスタを引いて、業務部所コードを取得して設定
                                    "LEFT JOIN Mst_Chousain mc ON " +
                                    "gjm.GyoumuJouhouMadoKojinCD = mc.KojinCD " +
                                    "LEFT JOIN Mst_Busho mb ON " + 
                                    "mc.GyoumuBushoCD = mb.GyoumuBushoCD " +

                                    "WHERE aj.AnkenJutakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + (string)w_rgnName[8, 2].Text + "' AND aj.AnkenDeleteFlag != 1 AND aj.AnkenSaishinFlg = 1 ";
                            var sda = new SqlDataAdapter(cmd);
                            sda.Fill(ankenJutakuDT);
                        }

                        string keiyakutantou = "0";
                        string ankenBangou = "";
                        string kanriGijutsushaCD = "";
                        string ankenJouhouID = "";
                        string jutakuBangou = "";
                        string jutakuBangouEda = "";
                        string kianzumi = "";
                        string ankenHachuushaKaMei = "";
                        string ankenGyoumuMei = "";
                        string GyoumuKanrishaCD = "";
                        string ankenKeiyakusho = "";
                        string gyoumuJouhouMadoGyoumuBushoCD = ""; // 業務部所CD
                        string gyoumuJouhouMadoKojinCD = "";       // 窓口担当者CD
                        string gyoumuJouhouMadoChousainMei = "";   // 窓口担当者名
                        string ankenJutakubushoCD = "";   // 受託課所支部
                        if (ankenJutakuDT != null && ankenJutakuDT.Rows.Count > 0)
                        {
                            // 契約担当者
                            keiyakutantou = ankenJutakuDT.Rows[0][0].ToString();
                            // 案件番号
                            ankenBangou = ankenJutakuDT.Rows[0][2].ToString();
                            // 管理技術者CD
                            kanriGijutsushaCD = ankenJutakuDT.Rows[0][3].ToString();
                            // AnkenJouhouID
                            ankenJouhouID = ankenJutakuDT.Rows[0][5].ToString();
                            // 受託番号
                            jutakuBangou = ankenJutakuDT.Rows[0][6].ToString();
                            // 受託番号枝番
                            jutakuBangouEda = ankenJutakuDT.Rows[0][7].ToString();
                            // 起案済み
                            kianzumi = ankenJutakuDT.Rows[0][8].ToString();
                            // 発注者名・課名
                            ankenHachuushaKaMei = ankenJutakuDT.Rows[0][9].ToString();
                            // 業務名称
                            ankenGyoumuMei = ankenJutakuDT.Rows[0][10].ToString();
                            // 業務管理者CD
                            GyoumuKanrishaCD = ankenJutakuDT.Rows[0][11].ToString();
                            // 案件フォルダ
                            ankenKeiyakusho = ankenJutakuDT.Rows[0][12].ToString();
                            // 業務部所CD
                            gyoumuJouhouMadoGyoumuBushoCD = ankenJutakuDT.Rows[0][13].ToString();
                            // 窓口担当者CD
                            gyoumuJouhouMadoKojinCD = ankenJutakuDT.Rows[0][14].ToString();
                            // 窓口担当者名
                            gyoumuJouhouMadoChousainMei = ankenJutakuDT.Rows[0][15].ToString();
                            // 受託課所支部
                            ankenJutakubushoCD = ankenJutakuDT.Rows[0][16].ToString();
                        }

                        tokuchoBangouEda = "";
                        ShukeiHyoFolder = "";
                        HoukokuShoFolder = "";
                        ShiryouHolder = "";

                        // 00:画面モード(insert)
                        chousagaiyouSQLData[0, 0] = "insert";
                        // このタイミングで採番してしまうとエラーがあった際に無駄にカウントアップしてしまうので、タイミングを変える
                        // 01:採番した窓口ID
                        //saibanMadoguchiID = GlobalMethod.getSaiban("MadoguchiId").ToString();
                        //chousagaiyouSQLData[0, 1] = saibanMadoguchiID;
                        // 02:登録年度
                        chousagaiyouSQLData[0, 2] = (string)w_rgnName[7, 2].Text;
                        if (int.TryParse(chousagaiyouSQLData[0, 2], out FromNendo))
                        {
                            ToNendo = FromNendo + 1;
                        }

                        // 03:遠隔地引渡承認]
                        chousagaiyouSQLData[0, 3] = "0";
                        // 04:遠隔地最終検査
                        chousagaiyouSQLData[0, 4] = "0";
                        // 05:遠隔地承認者
                        chousagaiyouSQLData[0, 5] = "";
                        // 06:遠隔地承認日
                        chousagaiyouSQLData[0, 6] = "null";
                        //// 07:調査担当者への締切日                 * (string)w_rgnName[22, 2].Text;
                        //chousagaiyouSQLData[0, 7] = getDate((string)w_rgnName[22, 2].Text);
                        ////if (chousagaiyouSQLData[0, 7] == "" || chousagaiyouSQLData[0, 7] == "null")
                        //if ((string)w_rgnName[22, 2].Text == "")
                        //{
                        //    // fileName, sheetName, row, msg
                        //    // E70060:必須入力項目が入力されていません。
                        //    //setErrorMsg(fileName, sheetName, 22, GlobalMethod.GetMessage("E70060", "調査担当者への締切日"));

                        //    workMsg += " 22行目:" + GlobalMethod.GetMessage("E70060", "調査担当者への締切日");
                        //    workFlg = true;

                        //    excelErrorFlg = true;
                        //}
                        //else if (chousagaiyouSQLData[0, 7] == "null")
                        //{
                        //    // E70067:調査担当者への締切日はYYYY/MM/DD形式で入力してください。
                        //    workMsg += " 22行目:" + GlobalMethod.GetMessage("E70067", "");
                        //    workFlg = true;

                        //    excelErrorFlg = true;
                        //}
                        // 08:登録日                               * (string)w_rgnName[20, 2].Text;
                        //chousagaiyouSQLData[0, 8] = getDate((string)w_rgnName[20, 2].Text);
                        chousagaiyouSQLData[0, 8] = getDate((string)w_rgnName[21, 2].Text);
                        //if (chousagaiyouSQLData[0, 8] == "" || chousagaiyouSQLData[0, 8] == "null")
                        if (chousagaiyouSQLData[0, 8] == "")
                        {
                            // fileName, sheetName, row, msg
                            // E70060:必須入力項目が入力されていません。
                            //setErrorMsg(fileName, sheetName, 20, GlobalMethod.GetMessage("E70060", "登録日"));

                            //workMsg += " 20行目:" + GlobalMethod.GetMessage("E70060", "登録日");
                            workMsg += " 21行目:" + GlobalMethod.GetMessage("E70060", "登録日");
                            workFlg = true;
                            excelErrorFlg = true;
                        }
                        else if(chousagaiyouSQLData[0, 8] == "null")
                        {
                            // E70068:登録日はYYYY/MM/DD形式で入力してください。
                            //workMsg += " 20行目:" + GlobalMethod.GetMessage("E70068", "");
                            workMsg += " 21行目:" + GlobalMethod.GetMessage("E70068", "");
                            workFlg = true;
                            excelErrorFlg = true;
                        }
                        // 07:調査担当者への締切日                 * (string)w_rgnName[22, 2].Text;
                        //chousagaiyouSQLData[0, 7] = getDate((string)w_rgnName[22, 2].Text);
                        chousagaiyouSQLData[0, 7] = getDate((string)w_rgnName[23, 2].Text);
                        //if (chousagaiyouSQLData[0, 7] == "" || chousagaiyouSQLData[0, 7] == "null")
                        //if ((string)w_rgnName[22, 2].Text == "")
                        if ((string)w_rgnName[23, 2].Text == "")
                        {
                            // fileName, sheetName, row, msg
                            // E70060:必須入力項目が入力されていません。
                            //setErrorMsg(fileName, sheetName, 22, GlobalMethod.GetMessage("E70060", "調査担当者への締切日"));

                            //workMsg += " 22行目:" + GlobalMethod.GetMessage("E70060", "調査担当者への締切日");
                            workMsg += " 23行目:" + GlobalMethod.GetMessage("E70060", "調査担当者への締切日");
                            workFlg = true;

                            excelErrorFlg = true;
                        }
                        else if (chousagaiyouSQLData[0, 7] == "null")
                        {
                            // E70067:調査担当者への締切日はYYYY/MM/DD形式で入力してください。
                            //workMsg += " 22行目:" + GlobalMethod.GetMessage("E70067", "");
                            workMsg += " 23行目:" + GlobalMethod.GetMessage("E70067", "");
                            workFlg = true;

                            excelErrorFlg = true;
                        }
                        // 09:報告実施日
                        //chousagaiyouSQLData[0, 9] = getDate((string)w_rgnName[24, 2].Text);
                        chousagaiyouSQLData[0, 9] = getDate((string)w_rgnName[25, 2].Text);
                        //if ((string)w_rgnName[24, 2].Text != "" && chousagaiyouSQLData[0, 9] == "null")
                        if ((string)w_rgnName[25, 2].Text != "" && chousagaiyouSQLData[0, 9] == "null")
                        {
                            // E70069:報告実施日はYYYY/MM/DD形式で入力してください。
                            //workMsg += " 24行目:" + GlobalMethod.GetMessage("E70069", "");
                            workMsg += " 25行目:" + GlobalMethod.GetMessage("E70069", "");
                            workFlg = true;

                            excelErrorFlg = true;
                        }

                        // 10:調査種別
                        //switch ((string)w_rgnName[16, 2].Text)
                        switch ((string)w_rgnName[17, 2].Text)
                        {
                            case "単品":
                                chousagaiyouSQLData[0, 10] = "1";
                                break;
                            case "一般":
                                chousagaiyouSQLData[0, 10] = "2";
                                break;
                            case "単契":
                                chousagaiyouSQLData[0, 10] = "3";
                                break;
                            default:
                                // 調査種別のデフォルトは1
                                chousagaiyouSQLData[0, 10] = "1";
                                break;
                        }
                        // 11:実施区分
                        //switch ((string)w_rgnName[18, 2].Text)
                        switch ((string)w_rgnName[19, 2].Text)
                        {
                            case "実施":
                                chousagaiyouSQLData[0, 11] = "1";
                                break;
                            case "診断中":
                                chousagaiyouSQLData[0, 11] = "2";
                                break;
                            case "中止":
                                chousagaiyouSQLData[0, 11] = "3";
                                break;
                            default:
                                // 実施区分のデフォルトは1
                                chousagaiyouSQLData[0, 11] = "1";
                                break;
                        }
                        // 12:MadoguchiShinchokuJoukyou 10:依頼
                        chousagaiyouSQLData[0, 12] = "10";
                        // 13:受託課所支部・・・プロンプトと特に連携していない項目で、デフォルトで自分の部所が入る
                        //chousagaiyouSQLData[0, 13] = UserInfos[2];
                        chousagaiyouSQLData[0, 13] = ankenJutakubushoCD;
                        // 14:契約担当者
                        chousagaiyouSQLData[0, 14] = keiyakutantou;
                        // 15:受託部所所属長の部所CD
                        chousagaiyouSQLData[0, 15] = UserInfos[2];
                        //// 16:窓口部所                             * (string)w_rgnName[5, 2].Text;
                        //if ((string)w_rgnName[5, 2].Text == "")
                        //{
                        //    // fileName, sheetName, row, msg
                        //    // E70060:必須入力項目が入力されていません。
                        //    //setErrorMsg(fileName, sheetName, 5, GlobalMethod.GetMessage("E70060","窓口部所"));

                        //    workMsg += " 5行目:" + GlobalMethod.GetMessage("E70060", "窓口部所");
                        //    workFlg = true;

                        //    excelErrorFlg = true;
                        //}
                        //else
                        //{
                        //    // 略名でくるので、CDを取得する
                        //    discript = "GyoumuBushoCD ";
                        //    value = "GyoumuBushoCD ";
                        //    table = "Mst_Busho ";
                        //    where = "BushoDeleteFlag != 1 AND BushokanriboKameiRaku = '" + (string)w_rgnName[5, 2].Text + "' ";

                        //    workDT = new DataTable();
                        //    workDT = GlobalMethod.getData(discript, value, table, where);
                        //    if (workDT != null && workDT.Rows.Count > 0)
                        //    {
                        //        chousagaiyouSQLData[0, 16] = workDT.Rows[0][0].ToString();
                        //    }
                        //    else
                        //    {
                        //        // fileName, sheetName, row, msg
                        //        // E20009:対象データは存在しません。
                        //        //setErrorMsg(fileName, sheetName, 5, GlobalMethod.GetMessage("E20009", "窓口部所"));

                        //        workMsg += " 5行目:" + GlobalMethod.GetMessage("E20009", "窓口部所");
                        //        workFlg = true;

                        //        excelErrorFlg = true;
                        //    }
                        //}
                        //// 17:窓口担当者
                        //if ((string)w_rgnName[6, 2].Text == "")
                        //{
                        //    chousagaiyouSQLData[0, 17] = "null";
                        //}
                        //else
                        //{
                        //    // 名称でくるので、CDを取得する
                        //    discript = "KojinCD ";
                        //    value = "KojinCD ";
                        //    table = "Mst_Chousain ";
                        //    where = "ChousainDeleteFlag != 1 AND ChousainMei = '" + (string)w_rgnName[6, 2].Text + "' ";

                        //    workDT = new DataTable();
                        //    workDT = GlobalMethod.getData(discript, value, table, where);
                        //    if (workDT != null && workDT.Rows.Count > 0)
                        //    {
                        //        chousagaiyouSQLData[0, 17] = workDT.Rows[0][0].ToString();
                        //    }
                        //}

                        // ①部所 無、担当者 無　⇒　エントリの窓口担当者で部所、担当者をセット
                        //⇒エントリの窓口担当者が取得できなかったらエラー
                        //（窓口担当者は必須なので、受託番号で案件が取れなかったらと同義）
                        // ②部所 有、担当者 無　⇒　エントリの窓口担当者で部所、担当者をセット
                        // ③部所 無、担当者 有　⇒　取込ファイルの担当者で部所と担当者をセット
                        // ④部所 有、担当者 有　⇒　取込ファイルの部所と担当者をセット

                        // 16:窓口部所                             * (string)w_rgnName[5, 2].Text;
                        if ((string)w_rgnName[5, 2].Text == "")
                        {
                            if (gyoumuJouhouMadoKojinCD == "")
                            {
                                // E70060:必須入力項目が入力されていません。
                                workMsg += " 5行目:" + GlobalMethod.GetMessage("E70060", "窓口部所");
                                workFlg = true;
                                excelErrorFlg = true;
                            }
                            else
                            {
                                chousagaiyouSQLData[0, 16] = gyoumuJouhouMadoGyoumuBushoCD;
                            }
                        }
                        else
                        {
                            // 略名でくるので、CDを取得する
                            discript = "GyoumuBushoCD ";
                            value = "GyoumuBushoCD ";
                            table = "Mst_Busho ";
                            where = "BushoDeleteFlag != 1 AND BushokanriboKameiRaku COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + (string)w_rgnName[5, 2].Text + "' ";
                            //where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                            //         "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
                            where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + DateTime.Today.ToString() + "' ) " +
                                     "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + DateTime.Today.ToString() + "' )";

                            workDT = new DataTable();
                            workDT = GlobalMethod.getData(discript, value, table, where);
                            if (workDT != null && workDT.Rows.Count > 0)
                            {
                                chousagaiyouSQLData[0, 16] = workDT.Rows[0][0].ToString();
                            }
                            else
                            {
                                // E20009:対象データは存在しません。
                                workMsg += " 5行目:" + GlobalMethod.GetMessage("E20009", "窓口部所");
                                workFlg = true;

                                excelErrorFlg = true;
                            }
                        }
                        // 17:窓口担当者(string)w_rgnName[6, 2].Text;
                        if ((string)w_rgnName[6, 2].Text == "")
                        {
                            //chousagaiyouSQLData[0, 17] = "null";
                            chousagaiyouSQLData[0, 17] = gyoumuJouhouMadoKojinCD;
                        }
                        else
                        {
                            // 名称でくるので、CDを取得する
                            discript = "GyoumuBushoCD ";
                            value = "KojinCD ";
                            table = "Mst_Chousain ";
                            where = "ChousainDeleteFlag != 1 AND ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + (string)w_rgnName[6, 2].Text + "' ";
                            //where += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                            //         "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' )";
                            where += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today.ToString() + "' ) " +
                                     "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today.ToString() + "' )";

                            workDT = new DataTable();
                            workDT = GlobalMethod.getData(discript, value, table, where);
                            if (workDT != null && workDT.Rows.Count > 0)
                            {
                                chousagaiyouSQLData[0, 17] = workDT.Rows[0][0].ToString();
                                // 担当者の部所と、窓口部所が異なっていた場合
                                if(workDT.Rows[0][1].ToString() != chousagaiyouSQLData[0, 16])
                                {
                                    chousagaiyouSQLData[0, 16] = workDT.Rows[0][1].ToString();
                                }
                            }
                            else
                            {
                                // E20329:担当者が存在しません。
                                workMsg += " 6行目:" + GlobalMethod.GetMessage("E20329", "窓口担当者");
                                workFlg = true;

                                excelErrorFlg = true;
                            }
                        }

                        // 18:窓口部所所属長の部所CD
                        chousagaiyouSQLData[0, 18] = chousagaiyouSQLData[0, 16];
                        // 19:調査区分　自部所
                        chousagaiyouSQLData[0, 19] = getAriNashi((string)w_rgnName[11, 2].Text);
                        // 20:調査区分 支→支
                        chousagaiyouSQLData[0, 20] = getAriNashi((string)w_rgnName[12, 2].Text);
                        // 21:調査区分 本→支
                        chousagaiyouSQLData[0, 21] = getAriNashi((string)w_rgnName[13, 2].Text);
                        // 22:調査区分 支→本
                        chousagaiyouSQLData[0, 22] = getAriNashi((string)w_rgnName[14, 2].Text);
                        // 調査区分は必須
                        if(chousagaiyouSQLData[0, 19] == "0" && chousagaiyouSQLData[0, 20] == "0" && chousagaiyouSQLData[0, 21] == "0" && chousagaiyouSQLData[0, 22] == "0")
                        {
                            // E70061:調査区分は、自部所、支→支、本→支、支→本のどれか1つ以上を有にしてください。
                            //setErrorMsg(fileName, sheetName, 5, GlobalMethod.GetMessage("E70061", ""));

                            workMsg += " 11,12,13,14行目:" + GlobalMethod.GetMessage("E70061", "");
                            workFlg = true;

                            excelErrorFlg = true;
                        }
                        //else if ((chousagaiyouSQLData[0, 20] == "1" && chousagaiyouSQLData[0, 21] == "1" && chousagaiyouSQLData[0, 22] == "1")
                        //    || (chousagaiyouSQLData[0, 20] == "1" && chousagaiyouSQLData[0, 21] == "1" && chousagaiyouSQLData[0, 22] == "0")
                        //    || (chousagaiyouSQLData[0, 20] == "1" && chousagaiyouSQLData[0, 21] == "0" && chousagaiyouSQLData[0, 22] == "1")
                        //    || (chousagaiyouSQLData[0, 20] == "0" && chousagaiyouSQLData[0, 21] == "1" && chousagaiyouSQLData[0, 22] == "1")
                        //    )
                        //{
                        //    // E70076:調査区分は自部所のみ、または自部所＋（支→支、本→支、支→本のどれか）、または（支→支、本→支、支→本のどれか）をチェックしてください。
                        //    workMsg += " 11,12,13,14行目:" + GlobalMethod.GetMessage("E70076", "");
                        //    workFlg = true;

                        //    excelErrorFlg = true;
                        //}

                        // 23:管理番号
                        chousagaiyouSQLData[0, 23] = (string)w_rgnName[10, 2].Text;
                        // 最大文字数まで取得
                        //len = 150; // 応援受付テーブル側が40文字
                        len = 40;
                        col = 23;
                        if(chousagaiyouSQLData[0, col].Length > len)
                        {
                            chousagaiyouSQLData[0, col] = chousagaiyouSQLData[0, col].Substring(0, len);
                        }

                        // 24:受託番号
                        chousagaiyouSQLData[0, 24] = (string)w_rgnName[8, 2].Text;
                        if (jutakuBangou == "")
                        {
                            // fileName, sheetName, row, msg
                            // E20009:対象データは存在しません。
                            //setErrorMsg(fileName, sheetName, 9, GlobalMethod.GetMessage("E20009", "受託番号"));

                            workMsg += " 8行目:" + GlobalMethod.GetMessage("E20009", "受託番号");
                            workFlg = true;

                            excelErrorFlg = true;
                        }
                        else
                        {
                            //// 起案済みでない場合
                            //if(kianzumi != "1")
                            //{
                            //    // E10008:起案されていないので、作成できません。
                            //    //setErrorMsg(fileName, sheetName, 9, GlobalMethod.GetMessage("E10008", "受託番号"));

                            //    workMsg += " 8行目:" + GlobalMethod.GetMessage("E10008", "受託番号");
                            //    workFlg = true;

                            //    excelErrorFlg = true;
                            //}

                        }
                        // 25:受託番号枝番
                        chousagaiyouSQLData[0, 25] = jutakuBangouEda;

                        // 26:特調番号
                        //chousagaiyouSQLData[0, 26] = ankenBangou;
                        // 2021/10/20
                        // 移行データの場合、案件番号にXが付く場合があり、受託番号と一致しないようので、
                        // 受託番号から-付き枝番をreplaceする方法に切り替える
                        chousagaiyouSQLData[0, 26] = jutakuBangou.Replace("-" + jutakuBangouEda,"");

                        // ファイル名の最後に_特調番号のチェックを行い、エラーとする
                        string fileName2 = fileName;
                        //if(fileName2 != "")
                        //{
                        //    // xxx_特調番号.xlsx から 特調番号.xlsx を取り出す
                        //    //fileName2 = fileName2.Substring(fileName2.LastIndexOf("_") + 1);

                        //    fileName2 = fileName2.Replace("【帳票フォーム】80_窓口情報取込_","");
                        //    fileName2 = fileName2.Replace(".xlsx", "");
                        //    fileName2 = fileName2.Replace(".xlsm", "");
                        //}
                        // 特調番号-枝番 とファイルの特調番号を比較する
                        //string tokucho = ankenBangou + "-" + (string)w_rgnName[9, 2].Text;
                        string tokucho = jutakuBangou.Replace("-" + jutakuBangouEda, "") + "-" + (string)w_rgnName[9, 2].Text;
                        //if (tokucho != fileName2)
                        // ファイル名に特調番号が存在するかどうか
                        //if (fileName2.IndexOf(tokucho) <= -1)
                        if (fileName2.IndexOf("_") != fileName2.IndexOf(tokucho) - 1)
                        {
                            // E70077:窓口ミハル一括登録ファイルの特調番号が一致しませんでした。
                            workMsg += " 9行目:" + GlobalMethod.GetMessage("E70077", "特調番号 枝番");
                            workFlg = true;

                            excelErrorFlg = true;
                        }


                        // 27:特調番号枝番
                        if ((string)w_rgnName[9, 2].Text == "")
                        {
                            // fileName, sheetName, row, msg
                            // E70060:必須入力項目が入力されていません。
                            //setErrorMsg(fileName, sheetName, 9, GlobalMethod.GetMessage("E70060", "特調番号 枝番"));

                            workMsg += " 9行目:" + GlobalMethod.GetMessage("E70060", "特調番号 枝番");
                            workFlg = true;

                            excelErrorFlg = true;
                        }
                        else
                        {
                            chousagaiyouSQLData[0, 27] = (string)w_rgnName[9, 2].Text;
                            tokuchoBangouEda = (string)w_rgnName[9, 2].Text;

                            // 重複チェック
                            // 特調番号・枝番号で被りがないか確認
                            workDT = new DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "MadoguchiID " +
                              "FROM MadoguchiJouhou " +
                              //"WHERE MadoguchiUketsukeBangou = '" + ankenBangou + "' " +
                              "WHERE MadoguchiUketsukeBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + jutakuBangou.Replace("-" + jutakuBangouEda, "") + "' " +
                              "AND MadoguchiUketsukeBangouEdaban COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + (string)w_rgnName[9, 2].Text + "' AND MadoguchiDeleteFlag != 1";

                            //データ取得
                            SqlDataAdapter tokuchoSda = new SqlDataAdapter(cmd);
                            tokuchoSda.Fill(workDT);

                            //データ取得できた場合
                            if (workDT.Rows.Count > 0)
                            {
                                //メッセージE20103を表示「特調番号が重複しました。」
                                //setErrorMsg(fileName, sheetName, 9, GlobalMethod.GetMessage("E20103", ""));

                                workMsg += " 9行目:" + GlobalMethod.GetMessage("E20103", "");
                                workFlg = true;

                                excelErrorFlg = true;
                            }
                        }
                        // 特調番号
                        //tokuchoBangou = ankenBangou + "-" + chousagaiyouSQLData[0, 27];
                        tokuchoBangou = jutakuBangou.Replace("-" + jutakuBangouEda, "") + "-" + chousagaiyouSQLData[0, 27];
                        // 28:発注者名・課名
                        //chousagaiyouSQLData[0, 28] = ankenHachuushaKaMei;
                        chousagaiyouSQLData[0, 28] = (string)w_rgnName[15, 2].Text;
                        // SE 20220309 No.1286 桁あふれ対処漏れの修正 
                        // 最大文字数まで取得                                                                // ADD 20220309
                        len = 150;                                                                          // ADD 20220309
                        col = 28;                                                                           // ADD 20220309
                        if (chousagaiyouSQLData[0, col].Length > len)                                       // ADD 20220309
                        {                                                                                   // ADD 20220309
                            chousagaiyouSQLData[0, col] = chousagaiyouSQLData[0, col].Substring(0, len);    // ADD 20220309
                        }                                                                                   // ADD 20220309
                        if(chousagaiyouSQLData[0, 28] == "")
                        {
                            chousagaiyouSQLData[0, 28] = ankenHachuushaKaMei;
                        }
                        // 29:業務名称
                        chousagaiyouSQLData[0, 29] = ankenGyoumuMei;
                        // 30:工事件名
                        //chousagaiyouSQLData[0, 30] = (string)w_rgnName[15, 2].Text;
                        chousagaiyouSQLData[0, 30] = (string)w_rgnName[16, 2].Text;
                        // 最大文字数まで取得
                        len = 150;
                        col = 30;
                        if (chousagaiyouSQLData[0, col].Length > len)
                        {
                            chousagaiyouSQLData[0, col] = chousagaiyouSQLData[0, col].Substring(0, len);
                        }
                        // 31:調査品目
                        //chousagaiyouSQLData[0, 31] = (string)w_rgnName[17, 2].Text;
                        chousagaiyouSQLData[0, 31] = (string)w_rgnName[18, 2].Text;
                        chousagaiyouSQLData[0, 31] = chousagaiyouSQLData[0, 31].Replace("\n", "\r\n");
                        // 最大文字数まで取得
                        len = 2048;
                        col = 31;
                        if (chousagaiyouSQLData[0, col].Length > len)
                        {
                            chousagaiyouSQLData[0, col] = chousagaiyouSQLData[0, col].Substring(0, len);
                        }
                        // 32:備考
                        //chousagaiyouSQLData[0, 32] = (string)w_rgnName[19, 2].Text;
                        chousagaiyouSQLData[0, 32] = (string)w_rgnName[20, 2].Text;
                        chousagaiyouSQLData[0, 32] = chousagaiyouSQLData[0, 32].Replace("\n", "\r\n");
                        // 最大文字数まで取得
                        len = 1024;
                        col = 32;
                        if (chousagaiyouSQLData[0, col].Length > len)
                        {
                            chousagaiyouSQLData[0, col] = chousagaiyouSQLData[0, col].Substring(0, len);
                        }
                        // 33:単価適用地域
                        //chousagaiyouSQLData[0, 33] = (string)w_rgnName[21, 2].Text;
                        chousagaiyouSQLData[0, 33] = (string)w_rgnName[22, 2].Text;
                        chousagaiyouSQLData[0, 33] = chousagaiyouSQLData[0, 33].Replace("\n", "\r\n");
                        // 最大文字数まで取得
                        len = 100;
                        col = 33;
                        if (chousagaiyouSQLData[0, col].Length > len)
                        {
                            chousagaiyouSQLData[0, col] = chousagaiyouSQLData[0, col].Substring(0, len);
                        }
                        // 34:荷渡場所
                        //chousagaiyouSQLData[0, 34] = (string)w_rgnName[23, 2].Text;
                        chousagaiyouSQLData[0, 34] = (string)w_rgnName[24, 2].Text;
                        // 最大文字数まで取得
                        len = 150;
                        col = 34;
                        if (chousagaiyouSQLData[0, col].Length > len)
                        {
                            chousagaiyouSQLData[0, col] = chousagaiyouSQLData[0, col].Substring(0, len);
                        }
                        // 35:報告済
                        chousagaiyouSQLData[0, 35] = getAriNashi((string)w_rgnName[2, 2].Text);
                        // 36:管理技術者
                        chousagaiyouSQLData[0, 36] = kanriGijutsushaCD;
                        // 37:本部単品
                        chousagaiyouSQLData[0, 37] = getAriNashi((string)w_rgnName[3, 2].Text);
                        //// 38:集計表フォルダ
                        //chousagaiyouSQLData[0, 38] = "";
                        //// 39:報告書フォルダ
                        //chousagaiyouSQLData[0, 39] = "";
                        //// 40:調査資料フォルダ
                        //chousagaiyouSQLData[0, 40] = "";

                        // 38:集計表フォルダ
                        chousagaiyouSQLData[0, 38] = ankenKeiyakusho + @"\" + GlobalMethod.GetCommonValue1("ANKEN_BANGOU_FOLDER", "13");
                        ShukeiHyoFolder = ankenKeiyakusho + @"\" + GlobalMethod.GetCommonValue1("ANKEN_BANGOU_FOLDER", "13");
                        // 39:報告書フォルダ
                        chousagaiyouSQLData[0, 39] = ankenKeiyakusho + @"\" + GlobalMethod.GetCommonValue1("ANKEN_BANGOU_FOLDER", "14");
                        HoukokuShoFolder = ankenKeiyakusho + @"\" + GlobalMethod.GetCommonValue1("ANKEN_BANGOU_FOLDER", "14");
                        // 40:調査資料フォルダ
                        chousagaiyouSQLData[0, 40] = ankenKeiyakusho + @"\" + GlobalMethod.GetCommonValue1("ANKEN_BANGOU_FOLDER", "15");
                        ShiryouHolder = ankenKeiyakusho + @"\" + GlobalMethod.GetCommonValue1("ANKEN_BANGOU_FOLDER", "15");

                        // 41:業務管理者の業務管理者CD
                        //chousagaiyouSQLData[0, 41] = "null";
                        chousagaiyouSQLData[0, 41] = GyoumuKanrishaCD;
                        // 42:AnkenJouhou.AnkenJouhouID
                        chousagaiyouSQLData[0, 42] = ankenJouhouID;
                        // 43:MadoguchiGaroonRenkei
                        chousagaiyouSQLData[0, 43] = getAriNashi((string)w_rgnName[4, 2].Text);
                        // 44:AnkenJouhou.AnkenJouhouID
                        chousagaiyouSQLData[0, 44] = ankenJouhouID;
                        // 45:連番
                        //chousagaiyouSQLData[0, 45] = TokuchoNo_renban(ankenBangou);
                        chousagaiyouSQLData[0, 45] = "0"; // 989:窓口ミハル一括登録中、途中で画面から操作が出来てしまう対応　更新タイミングを最後に持っていく

                        // 調査概要のエラー書き出し
                        if (workFlg == true)
                        {
                            setErrorMsg(fileName, sheetName, 1, workMsg);
                        }

                        // 担当部所

                        // 担当部所シート
                        if (sheetFlg2 == 0)
                        {
                            // ▼担当部所
                            sheetName = "担当部所";
                            worksheet = wb.Sheets[sheetName];
                            worksheet.Select();

                            workMsg = "";
                            workFlg = false;

                            xlSheet = null;
                            xlSheet = wb.Sheets[sheetName];
                            ws = wb.Sheets[sheetName];
                            //w_rgnName = xlSheet.UsedRange;
                            w_rgnName = xlSheet.UsedRange;

                            // Excel操作用の配列の作成
                            // 二次元配列の各次元の最小要素番号
                            int[] tantouLower = { 1, 2 };
                            // 二次元配列の各次元の要素数 
                            int[] tantouLength = { 1000, 2 };
                            object[,] tantouInputObject = (object[,])Array.CreateInstance(typeof(object), tantouLength, tantouLower);

                            // 使用されているエクセルの行数
                            getRowCount = ws.UsedRange.Rows.Count;

                            // 【担当部所】
                            //   共通ロジック                               Excelデータ
                            //-----------------------------------------------------------------------------------
                            //   00:ID
                            //   01:部所CD
                            // o 02:部所名                               (string)w_rgnName[i, 1].Text;
                            //   03:担当者CD
                            // o 04:担当者名                             (string)w_rgnName[i, 2].Text;

                            // A1から、B1000を取得する（宛先担当部所 宛先担当者）
                            Excel.Range tantouInputRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1000, 2]];
                            tantouInputObject = (object[,])tantouInputRange.Value;

                            string atesakiBushoCD = "";
                            string atesakiBusho = "";
                            string atesakiTantoushaCD = "";
                            string atesakiTantousha = "";

                            int garoonAtesakiRow = 0;
                            SqlDataAdapter tantouSda = null;
                            // Garoon追加宛先は1000行を範囲として見る 2行目のヘッダーまでを読み飛ばし
                            ////for (int i = 3; i < 1000; i++)
                            //for (int i = 3; i < getRowCount; i++)
                            for (int i = 4; i < getRowCount; i++)
                            {
                                // 宛先担当部所が空になったら処理終了
                                //if ((string)w_rgnName[i, 1] == null || (string)w_rgnName[i, 1].Text == "")
                                if (tantouInputObject[i, 1] == null || tantouInputObject[i, 1].ToString() == "")
                                {
                                    break;
                                }
                                garoonAtesakiRow += 1;

                                // 宛先担当部所
                                //atesakiBusho = (string)w_rgnName[i, 1].Text;
                                atesakiBusho = tantouInputObject[i, 1].ToString();

                                // 宛先担当者
                                //atesakiTantousha = (string)w_rgnName[i, 2].Text;
                                atesakiTantousha = tantouInputObject[i, 2].ToString();
                                if (atesakiTantousha == "")
                                {
                                    // E70060:必須入力項目が入力されていません。
                                    //setErrorMsg(fileName, sheetName, i, GlobalMethod.GetMessage("E70060", "宛先担当者"));

                                    workMsg += " " + i + "行目:" + GlobalMethod.GetMessage("E70060", "宛先担当者");
                                    workFlg = true;
                                    excelErrorFlg = true;
                                }
                                else
                                {
                                    // 宛先担当者が存在するか確認
                                    workDT = new DataTable();
                                    //SQL生成
                                    cmd.CommandText = "SELECT " +
                                      "KojinCD " +
                                      "FROM Mst_Chousain mc " +
                                      "LEFT JOIN Mst_Busho mb " +
                                      "  ON mc.GyoumuBushoCD = mb.GyoumuBushoCD " +
                                      "WHERE mc.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(atesakiTantousha, 0, 0) + "' " +
                                      "AND mb.BushokanriboKameiRaku COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(atesakiBusho, 0, 0) + "' ";
                                    cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today.ToString() + "' ) " +
                                             "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today.ToString() + "' ) ";
                                    cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + DateTime.Today.ToString() + "' ) " +
                                             "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + DateTime.Today.ToString() + "' )";

                                    //データ取得
                                    tantouSda = new SqlDataAdapter(cmd);
                                    tantouSda.Fill(workDT);

                                    // データ取得できた場合
                                    if (workDT != null && workDT.Rows.Count > 0)
                                    {
                                        // 正常
                                    }
                                    else
                                    {
                                        // E20329:担当者が存在しません。
                                        //setErrorMsg(fileName, sheetName, i, GlobalMethod.GetMessage("E20329", "宛先担当者"));

                                        workMsg += " " + i + "行目:" + GlobalMethod.GetMessage("E20329", "宛先担当者");
                                        workFlg = true;
                                        excelErrorFlg = true;
                                    }
                                }
                            }
                            // 存在した件数で配列を用意
                            tantoubushoGaroonAtesakiSQLData = new string[garoonAtesakiRow, 5];
                            for (int i = 4; i < 1000; i++)
                            {
                                // 宛先担当部所が空になったら処理終了
                                //if ((string)w_rgnName[i, 1].Text == "")
                                if (tantouInputObject[i, 1] == null || tantouInputObject[i, 1].ToString() == "")
                                {
                                    break;
                                }
                                // 宛先担当部所
                                //atesakiBusho = (string)w_rgnName[i, 1].Text;
                                atesakiBusho = tantouInputObject[i, 1].ToString();

                                // 宛先担当者
                                //atesakiTantousha = (string)w_rgnName[i, 2].Text;
                                atesakiTantousha = tantouInputObject[i, 2].ToString();

                                // 宛先担当者が存在するか確認
                                workDT = new DataTable();
                                //SQL生成
                                cmd.CommandText = "SELECT " +
                                  "mc.GyoumuBushoCD " +
                                  ",KojinCD " +
                                  "FROM Mst_Chousain mc " +
                                  "LEFT JOIN Mst_Busho mb " +
                                  "  ON mc.GyoumuBushoCD = mb.GyoumuBushoCD " +
                                  "WHERE mc.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(atesakiTantousha, 0, 0) + "' " +
                                  "AND mb.BushokanriboKameiRaku COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(atesakiBusho, 0, 0) + "' ";
                                cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today.ToString() + "' ) " +
                                         "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today.ToString() + "' )";
                                cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + DateTime.Today.ToString() + "' ) " +
                                         "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + DateTime.Today.ToString() + "' )";

                                //データ取得
                                workDT.Clear();
                                tantouSda = new SqlDataAdapter(cmd);
                                tantouSda.Fill(workDT);

                                // データ取得できた場合
                                if (workDT != null && workDT.Rows.Count > 0)
                                {
                                    // 正常
                                    atesakiBushoCD = workDT.Rows[0][0].ToString();
                                    atesakiTantoushaCD = workDT.Rows[0][1].ToString();
                                }

                                // 配列に詰める
                                // ID・・・未使用
                                tantoubushoGaroonAtesakiSQLData[i - 4, 0] = "";
                                tantoubushoGaroonAtesakiSQLData[i - 4, 1] = atesakiBushoCD;
                                tantoubushoGaroonAtesakiSQLData[i - 4, 2] = atesakiBusho;
                                tantoubushoGaroonAtesakiSQLData[i - 4, 3] = atesakiTantoushaCD;
                                tantoubushoGaroonAtesakiSQLData[i - 4, 4] = atesakiTantousha;
                            }

                            // 担当部所のエラー書き出し
                            if (workFlg == true)
                            {
                                setErrorMsg(fileName, sheetName, 2, workMsg);
                            }
                        }

                        // 調査品目明細




                        if (sheetFlg4 == 0)
                        {
                            // ▼協力依頼書
                            sheetName = "協力依頼書";
                            worksheet = wb.Sheets[sheetName];
                            worksheet.Select();

                            workMsg = "";
                            workFlg = false;

                            xlSheet = null;
                            xlSheet = wb.Sheets[sheetName];
                            ws = wb.Sheets[sheetName];
                            w_rgnName = xlSheet.UsedRange;

                            // Excel操作用の配列の作成
                            // 二次元配列の各次元の最小要素番号
                            int[] kyouryokuIraiLower = { 1, 1 };
                            // 二次元配列の各次元の要素数 
                            int[] kyouryokuIraiLength = { 15, 1 };
                            object[,] kyouryokuIraiInputObject = (object[,])Array.CreateInstance(typeof(object), kyouryokuIraiLength, kyouryokuIraiLower);

                            // B1から、B15を取得する
                            Excel.Range kyouryokuIraiRange = worksheet.Range[worksheet.Cells[1, 2], worksheet.Cells[15, 2]];
                            kyouryokuIraiInputObject = (object[,])kyouryokuIraiRange.Value;


                            kyouryokuIraishoID = "";
                            kyourokuIraisakiTantoshaCD = "";

                            kyouryokuIraiSda = null;
                            // 協力依頼書
                            // 【協力依頼書】
                            // 共通ロジック                             Excelデータ
                            //---------------------------------------------------------------------------
                            // 00:協力先部所                            (string)w_rgnName[2, 2].Text;
                            // 01:依頼日                                (string)w_rgnName[3, 2].Text;
                            // 02:報告期限
                            // 03:業務区分
                            // 04:依頼区分                              (string)w_rgnName[4, 2].Text;
                            // 05:内容区分 資材
                            // 06:内容区分 D工
                            // 07:内容区分 E工
                            // 08:内容区分 その他
                            // 09:内容区分 情報開発受託業務
                            // 10:連絡事項                              (string)w_rgnName[5, 2].Text;
                            // 11:業務内容                              (string)w_rgnName[6, 2].Text;
                            // 12:図面                                  (string)w_rgnName[7, 2].Text;
                            // 13:調査基準日（左のコンボ）              (string)w_rgnName[8, 2].Text;
                            // 14:調査基準日str                         (string)w_rgnName[9, 2].Text;
                            // 15:打合せ要否                            (string)w_rgnName[10, 2].Text;
                            // 16:具体的な協力先部課                    (string)w_rgnName[11, 2].Text;
                            // 17:前回協力（左のコンボ）                (string)w_rgnName[12, 2].Text;
                            // 18:前回協力str                           (string)w_rgnName[13, 2].Text;
                            // 19:成果物引渡場所                        (string)w_rgnName[14, 2].Text;
                            // 20:実施計画書                            (string)w_rgnName[15, 2].Text;
                            // 21:見積徴収                              (string)w_rgnName[16, 2].Text;
                            // 22:協力依頼書ID
                            // 23:協力先所属長

                            // 00:協力先部所kyouryokuIraishoSQLData 
                            // 全角スペースの場合、n
                            if ((string)w_rgnName[2, 2].Text != "")
                            {
                                // 支部名でくるので、CDを取得する
                                // 支部名でくるので、ShibuMeiを取得する
                                discript = "GyoumuBushoCD ";
                                value = "ShibuMei ";
                                table = "Mst_Busho ";
                                where = "BushoDeleteFlag != 1 AND ShibuMei COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + (string)w_rgnName[2, 2].Text + "' " +
                                // + "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                                //         "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' )";
                                "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "' ) " +
                                        "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "' )";

                                workDT = new DataTable();
                                workDT = GlobalMethod.getData(discript, value, table, where);
                                if (workDT != null && workDT.Rows.Count > 0)
                                {
                                    kyouryokuIraishoSQLData[0, 0] = workDT.Rows[0][0].ToString();
                                }
                                else
                                {
                                    // fileName, sheetName, row, msg
                                    // E20009:対象データは存在しません。
                                    //setErrorMsg(fileName, sheetName, 5, GlobalMethod.GetMessage("E20009", "協力先部所"));

                                    workMsg += " 5行目:" + GlobalMethod.GetMessage("E20009", "協力先部所");
                                    workFlg = true;
                                    excelErrorFlg = true;
                                }
                            }
                            else
                            {
                                kyouryokuIraishoSQLData[0, 0] = "null";
                            }
                            // 01:依頼日
                            kyouryokuIraishoSQLData[0, 1] = getDate((string)w_rgnName[3, 2].Text, 1);
                            if ((string)w_rgnName[3, 2].Text != "" && kyouryokuIraishoSQLData[0, 1] == "null")
                            {
                                // E70070:依頼日はYYYY/MM/DD形式で入力してください。
                                workMsg += " 3行目:" + GlobalMethod.GetMessage("E70070", "");
                                workFlg = true;

                                excelErrorFlg = true;
                            }
                            else
                            {
                                // nullの場合は、空文字をセット
                                if (kyouryokuIraishoSQLData[0, 1] == "null")
                                {
                                    kyouryokuIraishoSQLData[0, 1] = "";
                                }
                            }

                            // 02:報告期限
                            //kyouryokuIraishoSQLData[0, 2] = "";
                            kyouryokuIraishoSQLData[0, 2] = chousagaiyouSQLData[0, 7].Replace("'",""); // 調査概要の調査担当者への締切日がセットされる
                            // 03:業務区分
                            kyouryokuIraishoSQLData[0, 3] = "1"; // 1.一般受託調査
                            // 04:依頼区分
                            switch ((string)w_rgnName[4, 2].Text)
                            {
                                case "新規(契約前)":
                                    kyouryokuIraishoSQLData[0, 4] = "1";
                                    break;
                                case "新規(契約後)":
                                    kyouryokuIraishoSQLData[0, 4] = "2";
                                    break;
                                case "契約変更":
                                    kyouryokuIraishoSQLData[0, 4] = "3";
                                    break;
                                case "繰り返し協力依頼":
                                    kyouryokuIraishoSQLData[0, 4] = "4";
                                    break;
                                default:
                                    kyouryokuIraishoSQLData[0, 4] = "1";
                                    break;
                            }
                            // 05:内容区分 資材
                            kyouryokuIraishoSQLData[0, 5] = "0";
                            // 06:内容区分 D工
                            kyouryokuIraishoSQLData[0, 6] = "0";
                            // 07:内容区分 E工
                            kyouryokuIraishoSQLData[0, 7] = "0";
                            // 08:内容区分 その他
                            kyouryokuIraishoSQLData[0, 8] = "0";
                            // 09:内容区分 情報開発受託業務
                            kyouryokuIraishoSQLData[0, 9] = "0";
                            // 10:連絡事項
                            kyouryokuIraishoSQLData[0, 10] = (string)w_rgnName[5, 2].Text;
                            kyouryokuIraishoSQLData[0, 10] = kyouryokuIraishoSQLData[0, 10].Replace("\n", "\r\n");
                            // 最大文字数まで取得
                            len = 1024;
                            col = 10;
                            if (kyouryokuIraishoSQLData[0, col].Length > len)
                            {
                                kyouryokuIraishoSQLData[0, col] = kyouryokuIraishoSQLData[0, col].Substring(0, len);
                            }
                            // 11:業務内容
                            kyouryokuIraishoSQLData[0, 11] = (string)w_rgnName[6, 2].Text;
                            kyouryokuIraishoSQLData[0, 11] = kyouryokuIraishoSQLData[0, 11].Replace("\n", "\r\n");
                            // 最大文字数まで取得
                            len = 150;
                            col = 11;
                            if (kyouryokuIraishoSQLData[0, col].Length > len)
                            {
                                kyouryokuIraishoSQLData[0, col] = kyouryokuIraishoSQLData[0, col].Substring(0, len);
                            }
                            // 12:図面
                            //kyouryokuIraishoSQLData[0, 12] = getAriNashi((string)w_rgnName[7, 2].Text);
                            switch ((string)w_rgnName[7, 2].Text)
                            {
                                case "有":
                                    kyouryokuIraishoSQLData[0, 12] = "1";
                                    break;
                                case "無":
                                    kyouryokuIraishoSQLData[0, 12] = "2";
                                    break;
                                default:
                                    //kyouryokuIraishoSQLData[0, 12] = "0";
                                    kyouryokuIraishoSQLData[0, 12] = "2";
                                    break;
                            }

                            // 13:調査基準日（左のコンボ）
                            switch ((string)w_rgnName[8, 2].Text)
                            {
                                case "建設物価":
                                    kyouryokuIraishoSQLData[0, 13] = "1";
                                    break;
                                case "その他":
                                    kyouryokuIraishoSQLData[0, 13] = "2";
                                    break;
                                default:
                                    kyouryokuIraishoSQLData[0, 13] = "1";
                                    break;
                            }
                            // 14:調査基準日str
                            kyouryokuIraishoSQLData[0, 14] = w_rgnName[9, 2].Text;
                            // 最大文字数まで取得
                            len = 16;
                            col = 14;
                            if (kyouryokuIraishoSQLData[0, col].Length > len)
                            {
                                kyouryokuIraishoSQLData[0, col] = kyouryokuIraishoSQLData[0, col].Substring(0, len);
                            }
                            // 15:打合せ要否
                            switch ((string)w_rgnName[10, 2].Text)
                            {
                                case "否":
                                    kyouryokuIraishoSQLData[0, 15] = "2";
                                    break;
                                case "要":
                                    kyouryokuIraishoSQLData[0, 15] = "1";
                                    break;
                                default:
                                    kyouryokuIraishoSQLData[0, 15] = "0";
                                    break;
                            }
                            // 16:具体的な協力先部課
                            switch ((string)w_rgnName[11, 2].Text)
                            {
                                case "済":
                                    kyouryokuIraishoSQLData[0, 16] = "1";
                                    break;
                                default:
                                    kyouryokuIraishoSQLData[0, 16] = "0";
                                    break;
                            }
                            // 17:前回協力（左のコンボ）
                            switch ((string)w_rgnName[12, 2].Text)
                            {
                                case "有":
                                    kyouryokuIraishoSQLData[0, 17] = "1";
                                    break;
                                case "無":
                                    kyouryokuIraishoSQLData[0, 17] = "2";
                                    break;
                                default:
                                    //kyouryokuIraishoSQLData[0, 17] = "0";
                                    kyouryokuIraishoSQLData[0, 17] = "2";
                                    break;
                            }
                            // 18:前回協力str
                            kyouryokuIraishoSQLData[0, 18] = (string)w_rgnName[13, 2].Text;
                            // 最大文字数まで取得
                            len = 50;
                            col = 18;
                            if (kyouryokuIraishoSQLData[0, col].Length > len)
                            {
                                kyouryokuIraishoSQLData[0, col] = kyouryokuIraishoSQLData[0, col].Substring(0, len);
                            }
                            // 19:成果物引渡場所
                            switch ((string)w_rgnName[14, 2].Text)
                            {
                                case "協力先":
                                    kyouryokuIraishoSQLData[0, 19] = "1";
                                    break;
                                case "受託元":
                                    kyouryokuIraishoSQLData[0, 19] = "2";
                                    break;
                                default:
                                    kyouryokuIraishoSQLData[0, 19] = "1";
                                    break;
                            }
                            // 20:実施計画書
                            switch ((string)w_rgnName[15, 2].Text)
                            {
                                case "受託元で作成":
                                    kyouryokuIraishoSQLData[0, 20] = "1";
                                    break;
                                case "協力先で作成":
                                    kyouryokuIraishoSQLData[0, 20] = "2";
                                    break;
                                case "両部所で作成":
                                    kyouryokuIraishoSQLData[0, 20] = "3";
                                    break;
                                default:
                                    kyouryokuIraishoSQLData[0, 20] = "1";
                                    break;
                            }
                            // 21:見積徴収
                            switch ((string)w_rgnName[16, 2].Text)
                            {
                                case "協力先":
                                    kyouryokuIraishoSQLData[0, 21] = "1";
                                    break;
                                case "受託元":
                                    kyouryokuIraishoSQLData[0, 21] = "2";
                                    break;
                                case "両部所":
                                    kyouryokuIraishoSQLData[0, 21] = "3";
                                    break;
                                default:
                                    kyouryokuIraishoSQLData[0, 21] = "0";
                                    break;
                            }


                            // 協力依頼書のエラー書き出し
                            if (workFlg == true)
                            {
                                setErrorMsg(fileName, sheetName, 4, workMsg);
                            }
                        }
                        // 調査概要の登録が完了するまで取得できないのでタイミングを変える
                        //// 22:協力依頼書ID
                        //// 23:協力先所属長
                        //workDT = new DataTable();
                        ////SQL生成
                        //cmd.CommandText = "SELECT " +
                        //  "KyouryokuIraishoID " +
                        //  ",KyourokuIraisakiTantoshaCD " +
                        //  "FROM KyouryokuIraisho " +
                        //  "WHERE MadoguchiID = '" + saibanMadoguchiID + "' AND KyouryokuDeleteFlag <> 1";

                        ////データ取得
                        //workDT.Clear();
                        //kyouryokuIraiSda = new SqlDataAdapter(cmd);
                        //kyouryokuIraiSda.Fill(workDT);

                        //// データ取得できた場合
                        //if (workDT != null && workDT.Rows.Count > 0)
                        //{
                        //    // 正常
                        //    kyouryokuIraishoID = workDT.Rows[0][0].ToString();
                        //    kyourokuIraisakiTantoshaCD = workDT.Rows[0][1].ToString();
                        //}
                        //kyouryokuIraishoSQLData[0, 22] = kyouryokuIraishoID;
                        //kyouryokuIraishoSQLData[0, 23] = kyourokuIraisakiTantoshaCD;

                        if (sheetFlg5 == 0)
                        {
                            // ▼応援受付状況
                            sheetName = "応援受付状況";
                            worksheet = wb.Sheets[sheetName];

                            // 応援受付フラグ true：応援受付あり false:なし
                            ouenUketsukeFlg = true;

                            try
                            {
                                // シートが非表示だとここでエラーとなる
                                worksheet.Select();

                                workMsg = "";
                                workFlg = false;

                                xlSheet = null;
                                xlSheet = wb.Sheets[sheetName];
                                ws = wb.Sheets[sheetName];
                                w_rgnName = xlSheet.UsedRange;

                                // 【応援受付状況】
                                // 共通ロジック                             Excelデータ
                                //---------------------------------------------------------------------------
                                // 00:応援状況                            (string)w_rgnName[2, 2].Text;
                                // 01:応援受付日                          (string)w_rgnName[3, 2].Text;
                                // 02:応援完了                            (string)w_rgnName[4, 2].Text;
                                // 03:応援完了日                          (string)w_rgnName[5, 2].Text;

                                // 00:応援状況
                                ouenUketsukeSQLData[0, 0] = getAriNashi((string)w_rgnName[2, 2].Text);
                                // 応援状況は、1ではなく、2にしないといけない
                                if (ouenUketsukeSQLData[0, 0] == "1")
                                {
                                    ouenUketsukeSQLData[0, 0] = "2";
                                }
                                else
                                {
                                    // 支→本にチェックが入っていたら
                                    if(chousagaiyouSQLData[0, 22] == "1")
                                    {
                                        ouenUketsukeSQLData[0, 0] = "1";
                                    }
                                }

                                //if (getDate((string)w_rgnName[3, 2].Text, 1) != "null")
                                //{
                                //    ouenUketsukeSQLData[0, 1] = getDate((string)w_rgnName[3, 2].Text, 1);
                                //}
                                //else
                                //{
                                //    ouenUketsukeSQLData[0, 1] = "";
                                //}
                                ouenUketsukeSQLData[0, 1] = getDate((string)w_rgnName[3, 2].Text, 1);
                                if ((string)w_rgnName[3, 2].Text != "" && ouenUketsukeSQLData[0, 1] == "null")
                                {
                                    // E70071:応援受付日はYYYY/MM/DD形式で入力してください。
                                    workMsg += " 3行目:" + GlobalMethod.GetMessage("E70071", "");
                                    workFlg = true;

                                    excelErrorFlg = true;
                                }
                                else if((string)w_rgnName[3, 2].Text == "" && ouenUketsukeSQLData[0, 1] == "null")
                                {
                                    ouenUketsukeSQLData[0, 1] = "";
                                }

                                ouenUketsukeSQLData[0, 2] = getAriNashi((string)w_rgnName[4, 2].Text);
                                //if (getDate((string)w_rgnName[5, 2].Text, 1) != "null")
                                //{
                                //    ouenUketsukeSQLData[0, 3] = getDate((string)w_rgnName[5, 2].Text, 1);
                                //}
                                //else
                                //{
                                //    ouenUketsukeSQLData[0, 3] = "";
                                //}
                                ouenUketsukeSQLData[0, 3] = getDate((string)w_rgnName[5, 2].Text, 1);
                                if ((string)w_rgnName[5, 2].Text != "" && ouenUketsukeSQLData[0, 3] == "null")
                                {
                                    // E70072:応援完了日はYYYY/MM/DD形式で入力してください。
                                    workMsg += " 5行目:" + GlobalMethod.GetMessage("E70072", "");
                                    workFlg = true;

                                    excelErrorFlg = true;
                                }
                                else if ((string)w_rgnName[5, 2].Text == "" && ouenUketsukeSQLData[0, 3] == "null")
                                {
                                    ouenUketsukeSQLData[0, 3] = "";
                                }

                                // 応援受付状況のエラー書き出し
                                if (workFlg == true)
                                {
                                    setErrorMsg(fileName, sheetName, 5, workMsg);
                                }
                            }
                            catch (Exception)
                            {
                                // シート非表示
                                ouenUketsukeFlg = false;
                            }
                        }

                        if (sheetFlg6 == 0)
                        {
                            // ▼単品入力項目
                            sheetName = "単品入力項目";
                            worksheet = wb.Sheets[sheetName];
                            worksheet.Select();

                            workMsg = "";
                            workFlg = false;

                            xlSheet = null;
                            xlSheet = wb.Sheets[sheetName];
                            ws = wb.Sheets[sheetName];
                            w_rgnName = xlSheet.UsedRange;

                            // Excel操作用の配列の作成
                            // 二次元配列の各次元の最小要素番号
                            int[] tanpinLower = { 1, 1 };
                            // 二次元配列の各次元の要素数 
                            int[] tanpinLength = { 15, 1 };
                            object[,] tanpinInputObject = (object[,])Array.CreateInstance(typeof(object), tanpinLength, tanpinLower);

                            // B1から、B15を取得する
                            Excel.Range tanpinRange = worksheet.Range[worksheet.Cells[1, 2], worksheet.Cells[15, 2]];
                            tanpinInputObject = (object[,])tanpinRange.Value;

                            tanpinNyuuryokuID = "";

                            tanpinSda = null;
                            // 単品入力項目
                            // 【単品入力項目】
                            // 00.単品入力項目ID
                            // 01.受託日（依頼日）                     (string)w_rgnName[2, 2].Text;
                            // 02.報告日                               (string)w_rgnName[3, 2].Text;
                            // 03.指示番号                             (string)w_rgnName[4, 2].Text;
                            // 04.部署                                 (string)w_rgnName[6, 2].Text;
                            // 05.役職                                 (string)w_rgnName[7, 2].Text;
                            // 06.担当者                               (string)w_rgnName[8 2].Text;
                            // 07.電話                                 (string)w_rgnName[9, 2].Text;
                            // 08.FAX                                  (string)w_rgnName[10, 2].Text;
                            // 09.メール                               (string)w_rgnName[11, 2].Text;
                            // 10.メモ                                 (string)w_rgnName[12, 2].Text;
                            // 11.ランク                               未使用
                            // 12.照査実施                             (string)w_rgnName[14, 2].Text;
                            // 13.指示書                               (string)w_rgnName[16, 2].Text;
                            // 14.設計変更                             (string)w_rgnName[13, 2].Text;
                            // 15.見積提出方式                         (string)w_rgnName[15, 2].Text;
                            // 16.低入札                               (string)w_rgnName[17, 2].Text;
                            // 17.主要調査員                           未使用
                            // 18.単品請求月                           (string)w_rgnName[5, 2].Text;
                            // 19.市場価格（北陸専用）                　未使用
                            // 20.市場価格（北陸専用）r               　未使用
                            // 21.施工単価（北陸専用）                  未使用
                            // 22.施工単価（北陸専用）r                 未使用
                            // 23.その他集計
                            // 24.請求金額
                            // 25.請求確定
                            // 26.窓口ID
                            // 27.業務CD
                            // 28.契約情報ID                           未使用
                            // 29.経費（バックアップ用）                未使用

                            // 受託番号で単価契約から業務CDを取得する
                            DataTable TankaKeiyakuDT = new DataTable();
                            if (jutakuBangou != "")
                            {
                                //cmd.CommandText = "SELECT"
                                //                + " TankaKeiyakuID"
                                //                + " FROM TankaKeiyaku"
                                //                + " WHERE TankakeiyakuJutakuBangou = '" + jutakuBangou + "'"
                                //                ;

                                // 窓口の新規と合わせる
                                // ankenJouhouID は最新フラグが立っているもので引いている(TankaKeiyakuID の降順の一つ目を使用する)
                                cmd.CommandText = "SELECT TankaKeiyakuID FROM TankaKeiyaku"
                                                + " LEFT JOIN AnkenJouhou ON TankaKeiyaku.TankakeiyakuJutakuBangou = AnkenJouhou.AnkenJutakuBangou"
                                                + " WHERE AnkenJouhou.AnkenJouhouID = " + ankenJouhouID
                                                + " ORDER BY TankaKeiyakuID DESC ";

                                var sda = new SqlDataAdapter(cmd);
                                sda.Fill(TankaKeiyakuDT);
                            }

                            // ここでは取れないのでタイミングを変える
                            // 00.単品入力項目ID
                            //tanpinNyuuryokuSQLData[0, 0] = "";
                            // 01.受託日（依頼日）
                            //tanpinNyuuryokuSQLData[0, 1] = getDate((string)w_rgnName[2, 2].Text, 1);
                            //// nullの場合は、空文字をセット
                            //if (tanpinNyuuryokuSQLData[0, 1] == "null")
                            //{
                            //    tanpinNyuuryokuSQLData[0, 1] = "";
                            //}
                            tanpinNyuuryokuSQLData[0, 1] = getDate((string)w_rgnName[2, 2].Text, 1);
                            if ((string)w_rgnName[2, 2].Text != "" && tanpinNyuuryokuSQLData[0, 1] == "null")
                            {
                                // E70073:受託日（依頼日）はYYYY/MM/DD形式で入力してください。
                                workMsg += " 2行目:" + GlobalMethod.GetMessage("E70073", "");
                                workFlg = true;

                                excelErrorFlg = true;
                            }
                            else
                            {
                                if (tanpinNyuuryokuSQLData[0, 1] == "null")
                                {
                                    tanpinNyuuryokuSQLData[0, 1] = "";
                                }
                            }


                            // 02.報告日　契約報告日
                            //tanpinNyuuryokuSQLData[0, 2] = getDate((string)w_rgnName[3, 2].Text, 1);
                            //// nullの場合は、空文字をセット
                            //if (tanpinNyuuryokuSQLData[0, 2] == "null")
                            //{
                            //    tanpinNyuuryokuSQLData[0, 2] = "";
                            //}
                            tanpinNyuuryokuSQLData[0, 2] = getDate((string)w_rgnName[3, 2].Text, 1);
                            if ((string)w_rgnName[3, 2].Text != "" && tanpinNyuuryokuSQLData[0, 2] == "null")
                            {
                                // E70074:契約報告日はYYYY/MM/DD形式で入力してください。
                                workMsg += " 3行目:" + GlobalMethod.GetMessage("E70074", "");
                                workFlg = true;

                                excelErrorFlg = true;
                            }
                            else
                            {
                                if (tanpinNyuuryokuSQLData[0, 2] == "null")
                                {
                                    tanpinNyuuryokuSQLData[0, 2] = "";
                                }
                            }

                            // 03.指示番号
                            tanpinNyuuryokuSQLData[0, 3] = (string)w_rgnName[4, 2].Text;
                            // 最大文字数まで取得
                            len = 20;
                            col = 3;
                            if (tanpinNyuuryokuSQLData[0, col].Length > len)
                            {
                                tanpinNyuuryokuSQLData[0, col] = tanpinNyuuryokuSQLData[0, col].Substring(0, len);
                            }
                            // 04.部署
                            tanpinNyuuryokuSQLData[0, 4] = (string)w_rgnName[6, 2].Text;
                            tanpinNyuuryokuSQLData[0, 4] = tanpinNyuuryokuSQLData[0, 4].Replace("\n", "\r\n");
                            // 最大文字数まで取得
                            len = 150;
                            col = 4;
                            if (tanpinNyuuryokuSQLData[0, col].Length > len)
                            {
                                tanpinNyuuryokuSQLData[0, col] = tanpinNyuuryokuSQLData[0, col].Substring(0, len);
                            }
                            // 05.役職
                            tanpinNyuuryokuSQLData[0, 5] = (string)w_rgnName[7, 2].Text;
                            // 最大文字数まで取得
                            len = 20;
                            col = 5;
                            if (tanpinNyuuryokuSQLData[0, col].Length > len)
                            {
                                tanpinNyuuryokuSQLData[0, col] = tanpinNyuuryokuSQLData[0, col].Substring(0, len);
                            }
                            // 06.担当者
                            tanpinNyuuryokuSQLData[0, 6] = (string)w_rgnName[8, 2].Text;
                            // 最大文字数まで取得
                            len = 40;
                            col = 6;
                            if (tanpinNyuuryokuSQLData[0, col].Length > len)
                            {
                                tanpinNyuuryokuSQLData[0, col] = tanpinNyuuryokuSQLData[0, col].Substring(0, len);
                            }
                            // 07.電話
                            tanpinNyuuryokuSQLData[0, 7] = (string)w_rgnName[9, 2].Text;
                            // 最大文字数まで取得
                            len = 25;
                            col = 7;
                            if (tanpinNyuuryokuSQLData[0, col].Length > len)
                            {
                                tanpinNyuuryokuSQLData[0, col] = tanpinNyuuryokuSQLData[0, col].Substring(0, len);
                            }
                            if (!System.Text.RegularExpressions.Regex.IsMatch(tanpinNyuuryokuSQLData[0, col], @"^((0[0-9]{1,4}-[0-9]{1,4}-[0-9]{4})|\+?([0-9]{10,12})|(\+[0-9]{1,3}-[0-9]{1,2}-[0-9]{1,4}-[0-9]{4})|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                            {
                                // E20603:電話番号を正しく入力してください。
                                workMsg += " 9行目:" + GlobalMethod.GetMessage("E20603", "");
                                workFlg = true;
                                excelErrorFlg = true;
                            }
                            // 08.FAX
                            tanpinNyuuryokuSQLData[0, 8] = (string)w_rgnName[10, 2].Text;
                            // 最大文字数まで取得
                            len = 25;
                            col = 8;
                            if (tanpinNyuuryokuSQLData[0, col].Length > len)
                            {
                                tanpinNyuuryokuSQLData[0, col] = tanpinNyuuryokuSQLData[0, col].Substring(0, len);
                            }
                            if (!System.Text.RegularExpressions.Regex.IsMatch(tanpinNyuuryokuSQLData[0, col], @"^((0[0-9]{1,4}-[0-9]{1,4}-[0-9]{4})|\+?([0-9]{10,12})|(\+[0-9]{1,3}-[0-9]{1,2}-[0-9]{1,4}-[0-9]{4})|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                            {
                                // E20604:FAX番号を正しく入力してください。
                                workMsg += " 10行目:" + GlobalMethod.GetMessage("E20604", "");
                                workFlg = true;
                                excelErrorFlg = true;
                            }
                            // 09.メール
                            tanpinNyuuryokuSQLData[0, 9] = (string)w_rgnName[11, 2].Text;
                            // 最大文字数まで取得
                            len = 50;
                            col = 9;
                            if (tanpinNyuuryokuSQLData[0, col].Length > len)
                            {
                                tanpinNyuuryokuSQLData[0, col] = tanpinNyuuryokuSQLData[0, col].Substring(0, len);
                            }
                            if (!System.Text.RegularExpressions.Regex.IsMatch(tanpinNyuuryokuSQLData[0, col], @"^(([a-zA-Z_0-9]+([-+.'][a-zA-Z_0-9]+)*@[a-zA-Z_0-9]+([-.][a-zA-Z_0-9]+)*\.[a-zA-Z_0-9]+([-.][a-zA-Z_0-9]+)*)|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                            {
                                // E20605:メールアドレスを正しく入力してください。
                                workMsg += " 11行目:" + GlobalMethod.GetMessage("E20605", "");
                                workFlg = true;
                                excelErrorFlg = true;
                            }
                            // 10.メモ
                            tanpinNyuuryokuSQLData[0, 10] = (string)w_rgnName[12, 2].Text;
                            tanpinNyuuryokuSQLData[0, 10] = tanpinNyuuryokuSQLData[0, 10].Replace("\n", "\r\n");
                            // 最大文字数まで取得
                            len = 1024;
                            col = 10;
                            if (tanpinNyuuryokuSQLData[0, col].Length > len)
                            {
                                tanpinNyuuryokuSQLData[0, col] = tanpinNyuuryokuSQLData[0, col].Substring(0, len);
                            }
                            // 12.照査実施
                            tanpinNyuuryokuSQLData[0, 12] = (string)w_rgnName[14, 2].Text;
                            // 最大文字数まで取得
                            len = 10;
                            col = 12;
                            if (tanpinNyuuryokuSQLData[0, col].Length > len)
                            {
                                tanpinNyuuryokuSQLData[0, col] = tanpinNyuuryokuSQLData[0, col].Substring(0, len);
                            }
                            // 13.指示書
                            tanpinNyuuryokuSQLData[0, 13] = getAriNashi((string)w_rgnName[16, 2].Text);
                            // 14.設計変更
                            tanpinNyuuryokuSQLData[0, 14] = getAriNashi((string)w_rgnName[13, 2].Text);
                            // 15.見積提出方式
                            tanpinNyuuryokuSQLData[0, 15] = getAriNashi((string)w_rgnName[15, 2].Text);
                            // 16.低入札
                            tanpinNyuuryokuSQLData[0, 16] = getAriNashi((string)w_rgnName[17, 2].Text);
                            // 18.単品請求月
                            tanpinNyuuryokuSQLData[0, 18] = (string)w_rgnName[5, 2].Text;
                            // 最大文字数まで取得
                            len = 10;
                            col = 18;
                            if (tanpinNyuuryokuSQLData[0, col].Length > len)
                            {
                                tanpinNyuuryokuSQLData[0, col] = tanpinNyuuryokuSQLData[0, col].Substring(0, len);
                            }
                            // 23.その他集計 \0.00
                            tanpinNyuuryokuSQLData[0, 23] = "0";
                            // 24.請求金額 \0.00
                            tanpinNyuuryokuSQLData[0, 24] = "0";
                            // 25.請求確定
                            tanpinNyuuryokuSQLData[0, 25] = "0";

                            // 27.業務CD
                            tanpinNyuuryokuSQLData[0, 27] = "0";
                            if (TankaKeiyakuDT != null && TankaKeiyakuDT.Rows.Count > 0)
                            {
                                tanpinNyuuryokuSQLData[0, 27] = TankaKeiyakuDT.Rows[0][0].ToString();
                            }

                            // 単品入力項目のエラー書き出し
                            if (workFlg == true)
                            {
                                setErrorMsg(fileName, sheetName, 6, workMsg);
                            }

                        }
                        // ここでExcelを閉じておく
                        wb.Close(false, Type.Missing, Type.Missing);
                        // WorkBook Open フラグ 0:閉じている 1:開いている
                        workBookOpenFlg = 0;

                        // 正常の場合
                        if (excelErrorFlg == false)
                        {
                            // 01:採番した窓口ID
                            saibanMadoguchiID = GlobalMethod.getSaiban("MadoguchiId").ToString();
                            chousagaiyouSQLData[0, 1] = saibanMadoguchiID;

                            // 調査概要
                            // テーブル登録・更新
                            Boolean result = GlobalMethod.MadoguchiUpdate_SQL(1, saibanMadoguchiID, chousagaiyouSQLData, out resultMsg, UserInfos);

                            // 報告書フォルダ作成
                            // 窓口ID、特調番号の枝番
                            string resultMessage = GlobalMethod.CreateTokuchoBangouEdaFolder(saibanMadoguchiID, chousagaiyouSQLData[0, 27]);

                            // 担当部所
                            string[,] dummySQLData = new string[1, 2];
                            string mes = "";
                            if (sheetFlg2 == 0)
                            {
                                GlobalMethod.MadoguchiUpdate_SQL(2, saibanMadoguchiID, dummySQLData, out mes, UserInfos, tantoubushoGaroonAtesakiSQLData);
                            }

                            if (sheetFlg3 == 0)
                            {
                                // 調査品目明細
                                string[] results = GlobalMethod.InsertHinmoku(filePath, saibanMadoguchiID, UserInfos[0], UserInfos[2], "1");

                                chousahinmokuErrorId = 0;

                                // result
                                // 成否判定 0:正常 1：エラー
                                // T_ReadFileErrorテーブルのエラーカウント（FileReadErrorReadCount）
                                // メッセージ（主にエラー用）
                                if (results != null && results.Length >= 1)
                                {
                                    // 改行コードがあるので、削る
                                    results[0] = results[0].Replace(@"\r\n", "");

                                    if (results[0].Trim() == "1")
                                    {
                                        //// エラーが発生しました
                                        //set_error(GlobalMethod.GetMessage("E00091", ""));
                                        //set_error(results[2]);
                                        int count = 0;
                                        // T_ReadFileErrorテーブルのエラーカウントをセット
                                        if (results[1] != null && int.TryParse(results[1].ToString(), out count))
                                        {
                                            // 調査品目で採番されたエラーID
                                            chousahinmokuErrorId = count;

                                            // エラーが発生したので、窓口情報を消す
                                            madoguchiDelete();

                                            // GeneXusではファイル名をフルパスで登録するので、ファイル名だけに置き換える
                                            historyFileNameUpdate(fileName);

                                            registrationResult = false;
                                            // 後続の処理をしない為に処理を終了する
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        // 窓口連携
                                        GlobalMethod.MadoguchiHinmokuRenkeiUpdate_SQL(saibanMadoguchiID, "Madoguchi", UserInfos[0], out resultMessage);
                                    }
                                }
                            }

                            if (sheetFlg4 == 0)
                            {
                                // 22:協力依頼書ID
                                // 23:協力先所属長
                                workDT = new DataTable();
                                //SQL生成
                                cmd.CommandText = "SELECT " +
                                  "KyouryokuIraishoID " +
                                  ",KyourokuIraisakiTantoshaCD " +
                                  "FROM KyouryokuIraisho " +
                                  "WHERE MadoguchiID = '" + saibanMadoguchiID + "' AND KyouryokuDeleteFlag <> 1";

                                //データ取得
                                workDT.Clear();
                                kyouryokuIraiSda = new SqlDataAdapter(cmd);
                                kyouryokuIraiSda.Fill(workDT);

                                kyouryokuIraishoID = "0";
                                kyourokuIraisakiTantoshaCD = "null";

                                // データ取得できた場合
                                if (workDT != null && workDT.Rows.Count > 0)
                                {
                                    // 正常
                                    kyouryokuIraishoID = workDT.Rows[0][0].ToString();
                                    kyourokuIraisakiTantoshaCD = workDT.Rows[0][1].ToString();
                                }
                                kyouryokuIraishoSQLData[0, 22] = kyouryokuIraishoID;

                                workDT = new DataTable();
                                //SQL生成
                                cmd.CommandText = "SELECT " +
                                  "ChousainMei, KojinCD " +
                                  "FROM Mst_Busho " +
                                  "LEFT JOIN Mst_Chousain ON BushoShozokuChou = ChousainMei " +
                                  "WHERE Mst_Busho.ShibuMei COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + kyouryokuIraishoSQLData[0, 0] + "' ";
                                //cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                                //         "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' )";
                                cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today.ToString() + "' ) " +
                                         "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today.ToString() + "' )";
                                cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + DateTime.Today.ToString() + "' ) " +
                                         "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + DateTime.Today.ToString() + "' )";

                                //データ取得
                                workDT.Clear();
                                kyouryokuIraiSda = new SqlDataAdapter(cmd);
                                kyouryokuIraiSda.Fill(workDT);

                                // データ取得できた場合
                                if (workDT != null && workDT.Rows.Count > 0)
                                {
                                    // 正常
                                    kyourokuIraisakiTantoshaCD = workDT.Rows[0][1].ToString();
                                }

                                kyouryokuIraishoSQLData[0, 23] = kyourokuIraisakiTantoshaCD;

                                // 協力依頼書

                                GlobalMethod.MadoguchiUpdate_SQL(4, saibanMadoguchiID, kyouryokuIraishoSQLData, out mes, UserInfos);
                            }

                            // --施行条件・・・テンプレートから削除
                            if (sheetFlg5 == 0)
                            {
                                // シートがあれば更新
                                if (ouenUketsukeFlg == true)
                                {
                                    // 応援受付状況応援受付状況
                                    GlobalMethod.MadoguchiUpdate_SQL(5, saibanMadoguchiID, ouenUketsukeSQLData, out mes, UserInfos);
                                }
                            }

                            if (sheetFlg6 == 0)
                            {
                                // 単品入力項目
                                // 単品入力ID
                                workDT = new DataTable();
                                //SQL生成
                                cmd.CommandText = "SELECT " +
                                  "TanpinNyuuryokuID " +
                                  "FROM TanpinNyuuryoku " +
                                  "WHERE MadoguchiID = '" + saibanMadoguchiID + "' AND TanpinDeleteFlag <> 1";

                                //データ取得
                                workDT.Clear();
                                tanpinSda = new SqlDataAdapter(cmd);
                                tanpinSda.Fill(workDT);

                                tanpinNyuuryokuID = "0";

                                // データ取得できた場合
                                if (workDT != null && workDT.Rows.Count > 0)
                                {
                                    // 正常
                                    tanpinNyuuryokuID = workDT.Rows[0][0].ToString();
                                }
                                tanpinNyuuryokuSQLData[0, 0] = tanpinNyuuryokuID;

                                tanpinNyuuryokuSQLData[0, 26] = saibanMadoguchiID;

                                dummySQLData[0, 1] = null;
                                GlobalMethod.MadoguchiUpdate_SQL(6, saibanMadoguchiID, tanpinNyuuryokuSQLData, out mes, UserInfos, dummySQLData);
                            }

                            string message = "";
                            GlobalMethod.garoonRenkeiUpdate(saibanMadoguchiID, UserInfos, out message);

                            // 989:窓口ミハル一括登録中、途中で画面から操作が出来てしまう対応
                            //cmd.CommandText = "UPDATE MadoguchiJouhou SET MadoguchiSystemRenban = " + TokuchoNo_renban(ankenBangou) + " " +
                            cmd.CommandText = "UPDATE MadoguchiJouhou SET MadoguchiSystemRenban = " + TokuchoNo_renban(jutakuBangou.Replace("-" + jutakuBangouEda, "")) + " " +
                                              " WHERE MadoguchiID = '" + saibanMadoguchiID + "'";
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            // エラーチェックに引っかかった場合

                            // 調査品目明細のチェック or 登録は exe なので、ファイルにエラーがあった時点で調査品目明細のエラーチェック不可
                            // E70064:調査品目一覧シートについては、調査概要を登録できなかったので、チェックできませんでした。
                            setErrorMsg(fileName, "調査品目一覧", 3, GlobalMethod.GetMessage("E70064", ""));

                            registrationResult = false;
                        }

                        transaction.Commit();


                        // 協力依頼書の自動出力
                        if (excelErrorFlg == false && (chousagaiyouSQLData[0, 20] == "1" || chousagaiyouSQLData[0, 21] == "1" || chousagaiyouSQLData[0, 22] == "1"))
                        {
                            string MadoguchiShiryouHolder = "";
                            DataTable dt = new DataTable();
                            dt = GlobalMethod.getData("MadoguchiShiryouHolder", "MadoguchiShiryouHolder", "MadoguchiJouhou", "MadoguchiID = " + saibanMadoguchiID);

                            if(dt != null && dt.Rows.Count > 0)
                            {
                                MadoguchiShiryouHolder = dt.Rows[0][0].ToString();
                            }
                            // 資料フォルダが存在するか
                            if(MadoguchiShiryouHolder != "" && Directory.Exists(MadoguchiShiryouHolder))
                            {
                                string[] report_data = new string[2] { "", "" };

                                report_data[0] = saibanMadoguchiID;
                                report_data[1] = "1";

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
                                            System.IO.File.Copy(result[2], MadoguchiShiryouHolder + @"\" + copyFileName, true);

                                            workDT = new DataTable();
                                            //SQL生成
                                            cmd.CommandText = "SELECT " +
                                              "KyourokuIraisakiBushoOld,KyouryokuDate,KyouryokuIraiKubun,KyouryokuGyoumuNaiyou " +
                                              ",KyouryokuZumen,KyouryokuChousaKijun,KyouryokuUtiawaseyouhi " +
                                              ",KyouryokuZenkaiUmu,KyouryokusakiHikiwatashi,KyouryokuChoushuusaki " +
                                              "FROM KyouryokuIraisho with(nolock) " +
                                              "WHERE MadoguchiID = '" + saibanMadoguchiID + "' AND KyouryokuDeleteFlag <> 1";

                                            //データ取得
                                            workDT.Clear();
                                            SqlDataAdapter kyouSda = new SqlDataAdapter(cmd);
                                            kyouSda.Fill(workDT);

                                            if(workDT != null && workDT.Rows.Count > 0)
                                            {
                                                // 協力依頼書出力で必須チェックしている項目をチェックする
                                                if (workDT.Rows[0][0].ToString() == ""     // 協力先部所
                                                    || workDT.Rows[0][1].ToString() == ""  // 依頼日
                                                    || workDT.Rows[0][2].ToString() == ""  // 依頼区分
                                                    || workDT.Rows[0][3].ToString() == ""  // 業務内容
                                                    || workDT.Rows[0][4].ToString() == "0"  // 図面
                                                    || workDT.Rows[0][5].ToString() == ""  // 調査基準
                                                    || workDT.Rows[0][6].ToString() == "0"  // 打合せ要否
                                                    || workDT.Rows[0][7].ToString() == "0"  // 前回協力
                                                    || workDT.Rows[0][8].ToString() == ""  // 成果物引渡場所
                                                    || workDT.Rows[0][9].ToString() == "0"  // 見積徴収
                                                    )
                                                {
                                                    // W20202:自動出力した協力依頼書には一部未設定の項目があるので、必要に応じて画面から設定し、再出力してください。
                                                    set_error(GlobalMethod.GetMessage("W20202", ""));
                                                }
                                            }
                                        }
                                        catch (Exception)
                                        {
                                            // W20201:協力依頼書が出力できませんでした。協力依頼書タブから手動で出力してください。
                                            set_error(GlobalMethod.GetMessage("W20201",""));
                                        }
                                    }
                                }
                            }
                            // 存在していない場合、メッセージを表示
                            else
                            {
                                // W20201:協力依頼書が出力できませんでした。協力依頼書タブから手動で出力してください。
                                set_error(GlobalMethod.GetMessage("W20201", ""));
                            }
                        }
                    }
                    catch (Exception e2)
                    {
                        transaction.Rollback();

                        // 上記のエラーチェックを通ったが、GlobalMethod側でエラーとなった場合
                        // GlobalMethod でCommitしてしまっているので、ここで登録したデータを削除する

                        // ▼以下の順で窓口のデータを削除していく
                        // 支部備考欄テーブル（ShibuBikou）
                        // 施工条件テーブル（SekouJouken）
                        // 応援受付テーブル（OuenUketsuke）
                        // 協力依頼書情報テーブル（KyouryokuIraisho）
                        // 単品入力項目テーブル（TanpinNyuuryoku）
                        // 調査品目情報テーブル（ChousaHinmoku）
                        // 窓口情報（調査担当者）テーブル（MadoguchiJouhouMadoguchiL1Chou）
                        // 窓口情報テーブル（MadoguchiJouhou）
                        try
                        {
                            madoguchiDelete();

                            writeHistory("窓口ミハル一括取込で予期せぬエラーが有った為、登録した窓口情報を削除しました ID = " + saibanMadoguchiID);

                        }
                        catch(Exception e)
                        {

                        }
                        // ここでメッセージを出してもかき消される
                        //set_error("", 0);
                        //// E00091:エラーが発生しました。
                        //set_error(GlobalMethod.GetMessage("E00091", ""), 0);

                        // 例外発生フラグ 0:発生なし 1:発生
                        exceptionFlg = 1;

                        return false;
                    }
                    finally
                    {
                        // Rangeもプロセス開放しないといけない
                        Marshal.ReleaseComObject(w_rgnName);
                    }
                    conn.Close();
                }
               return registrationResult;
            }
            catch (Exception)
            {
                return false;
            }
            return registrationResult;
        }

        // 窓口削除
        private void madoguchiDelete()
        {
            if(saibanMadoguchiID != "")
            {
                try
                {
                    var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                    using (var conn = new SqlConnection(connStr))
                    {
                        //エラーメッセージ
                        conn.Open();
                        var cmd = conn.CreateCommand();

                        cmd.CommandText = "DELETE FROM ShibuBikou "
                                                + " WHERE MadoguchiID = '" + saibanMadoguchiID + "'"
                                                ;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "DELETE FROM SekouJouken "
                                        + " WHERE MadoguchiID = '" + saibanMadoguchiID + "'"
                                        ;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "DELETE FROM OuenUketsuke "
                                        + " WHERE MadoguchiID = '" + saibanMadoguchiID + "'"
                                        ;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "DELETE FROM KyouryokuIraisho "
                                        + " WHERE MadoguchiID = '" + saibanMadoguchiID + "'"
                                        ;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "DELETE FROM TanpinNyuuryoku "
                                        + " WHERE MadoguchiID = '" + saibanMadoguchiID + "'"
                                        ;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "DELETE FROM ChousaHinmoku "
                                        + " WHERE MadoguchiID = '" + saibanMadoguchiID + "'"
                                        ;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "DELETE FROM MadoguchiJouhouMadoguchiL1Chou "
                                        + " WHERE MadoguchiID = '" + saibanMadoguchiID + "'"
                                        ;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "DELETE FROM MailInfoCSVWork " 
                                        + " WHERE MailInfoCSVWorkMadoguchiID = '" + saibanMadoguchiID + "'"
                                        ;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "DELETE FROM GaroonTsuikaAtesaki "
                                        +  " WHERE GaroonTsuikaAtesakiMadoguchiID = '" + saibanMadoguchiID + "'"
                                        ;
                        cmd.ExecuteNonQuery();

                        conn.Close();


                        conn.Open();
                        cmd = conn.CreateCommand();
                        // T_HistoryのMadoguhiIDが外部キーついてて窓口情報が消せないので、Updateをかける
                        cmd.CommandText = "UPDATE T_HISTORY SET MadoguchiID = null "
                                        + " WHERE MadoguchiID = '" + saibanMadoguchiID + "'"
                                        ;
                        cmd.ExecuteNonQuery();


                        cmd.CommandText = "DELETE FROM MadoguchiJouhou "
                                        + " WHERE MadoguchiID = '" + saibanMadoguchiID + "'"
                                        ;
                        cmd.ExecuteNonQuery();

                        string MadoguchiShukeiHyoFolder = "";  // 集計表
                        string MadoguchiHoukokuShoFolder = ""; // 報告書
                        string MadoguchiShiryouHolder = "";    // 調査資料
                        string targetFolder = "";

                        MadoguchiShukeiHyoFolder = ShukeiHyoFolder + @"\" + tokuchoBangouEda;
                        MadoguchiHoukokuShoFolder = HoukokuShoFolder + @"\" + tokuchoBangouEda;
                        MadoguchiShiryouHolder = ShiryouHolder + @"\" + tokuchoBangouEda;

                        // 集計表
                        targetFolder = MadoguchiShukeiHyoFolder;
                        //  
                        if (Directory.Exists(targetFolder))
                        {
                            DirectoryInfo di = new DirectoryInfo(targetFolder);
                            di.Delete();
                        }
                        // 報告書
                        targetFolder = MadoguchiHoukokuShoFolder;
                        // フォルダの存在チェックをし、削除
                        if (Directory.Exists(targetFolder))
                        {
                            DirectoryInfo di = new DirectoryInfo(targetFolder);
                            di.Delete();
                        }
                        // 調査資料
                        targetFolder = MadoguchiShiryouHolder;
                        // フォルダの存在チェックをし、削除
                        if (Directory.Exists(targetFolder))
                        {
                            DirectoryInfo di = new DirectoryInfo(targetFolder);
                            di.Delete();
                        }

                        conn.Close();
                    }
                }
                catch (Exception dele)
                {
                    writeHistory("窓口情報削除処理でエラー：" + dele.Message);
                    throw;
                }
            }
        }

        private void historyFileNameUpdate(string fileName)
        {
            try
            {
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    //エラーメッセージ
                    conn.Open();
                    var cmd = conn.CreateCommand();

                    cmd.CommandText = "UPDATE T_FileReadError set FileReadErrorFilenameCOLLATE Japanese_XJIS_100_CI_AS_SC  = N'" + fileName + "' "
                                    + " WHERE FileReadErrorTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + saibanMadoguchiID + "'";
                    cmd.ExecuteNonQuery();

                    conn.Close();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private Workbook getExcelFile(String fileName, Application ExcelApp)
        {
            //Excelシートのインスタンスを作る
            if (!System.IO.File.Exists(fileName))
            {
                MessageBox.Show("'" + fileName + "'は存在しません。");
                return null;
            }

            Excel.Workbook wb = ExcelApp.Workbooks.Open(fileName);
            int num = 0;
            // wait秒数 デフォルト3秒
            string strWork = GlobalMethod.GetCommonValue1("MADOGUCHI_IKKATSU_WAIT_TIME");
            if (strWork != null)
            {
                if (int.TryParse(strWork, out num))
                {
                    waitTime = num;
                }
                else
                {
                    waitTime = 3;
                }
            }
            else
            {
                waitTime = 3;
            }

            num = 0;
            // 準備完了まで待つ
            while (true)
            {
                if (ExcelApp.Ready == true || waitTime <= num)
                {
                    break;
                }
                Thread.Sleep(1000);
                num += 1;
            }

            ExcelApp.Visible = false;
            return wb;
        }

        // VIPS　20220228　課題管理表No1259(949)　DEL　「窓口ミハル一括登録取り込み結果」ボタン非表示  対応
        //// 窓口ミハル一括取込結果一覧
        //private void BtnTorikomiKekka_Click(object sender, EventArgs e)
        //{
        //    set_error("", 0);
        //    if(BtnTorikomiKekka.BackColor == Color.FromArgb(42, 78, 122))
        //    {
        //        Popup_FileError form = new Popup_FileError(errorUser, readCount);
        //        // 調査品目のエラーがあれば
        //        if (chousahinmokuErrorId > 0)
        //        {
        //            form.chousaHinmokuErrorCnt = chousahinmokuErrorId.ToString();
        //            form.FileReadErrorTokuchoBangou = saibanMadoguchiID;
        //        }
        //        form.ShowDialog();
        //    }
        //}

        // バックグラウンドとなっているExcleのプロセスをKillする
        private void excelProcessKill()
        {
            try
            {
                // 他のExcelを起動していてもバックグラウンドとなっているプロセスのみをKILLするので影響はない、らしい
                foreach (var p in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    if (p.MainWindowTitle == "")
                    {
                        try
                        {
                            p.Kill();
                        }
                        // 既に終了 or 終了処理中 となっているプロセスをKILL仕様とした場合、スルー
                        catch (InvalidOperationException e)
                        {

                        }
                    }
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // E00090:システム障害が発生しました。
                //set_error(GlobalMethod.GetMessage("E00090", "EXCEL Process kill"));
            }
        }

        // 金額チェック（true：OK、false：NG）
        private Boolean CheckMoney(string strMoney)
        {
            if (Regex.IsMatch(strMoney, @"^[\-]*[0-9]+$"))
            {
                return true;
            }
            return false;
        }

        // 数字チェック（true：OK、false：NG）
        private Boolean CheckNumber(string strNumber)
        {
            if (Regex.IsMatch(strNumber, @"^[0-9]+$"))
            {
                return true;
            }
            return false;
        }

        // 特調番号のシステム連番を取得する
        private String TokuchoNo_renban(string uketsukeBangou)
        {
            string tokuchoNo = "1";

            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                var dt = new DataTable();
                cmd.CommandText = "SELECT ISNULL(MAX(MadoguchiSystemRenban),0) +1 AS renbanMax FROM MadoguchiJouhou " +
                "WHERE MadoguchiUketsukeBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + uketsukeBangou + "' " +
                "AND MadoguchiDeleteFlag != 1 ";

                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);

                //連番を取得できたらそれをセットする
                tokuchoNo = dt.Rows[0][0].ToString();

            }

            return tokuchoNo;
        }

        // 有は1,それ以外は0を返す
        private String getAriNashi(string str)
        {
            string rtnStr = "0";
            if(str == "有")
            {
                rtnStr = "1";
            }
            return rtnStr;
        }

        // YYYY/MM/DD 形式のデータを返す 形式に合わない場合、空文字を返す
        private String getDate(string str, int flg = 0)
        {
            string rtnStr = "null";
            DateTime dt;
            // TryParseで日付変換出来るかを調べる
            if (DateTime.TryParse(str, out dt))
            {
                // 文字列に変換
                rtnStr = dt.ToString("yyyy/MM/dd");
                // flg 0:'で囲む それ以外囲まない
                if (flg == 0)
                {
                    // ' で囲む
                    rtnStr = "'" + rtnStr + "'";
                }
            }

            return rtnStr;
        }

        // ErrorMsg、errorId、readCount をセットする
        private void setErrorMsg(string fileName, string sheetName, int row, string msg)
        {
            // メッセージ初期化
            ErrorMsg = "";
            if (ErrorMsg == "")
            {
                ErrorMsg = sheetName + "シート ";
            }
            if (errorId == 0)
            {
                errorId = GlobalMethod.getSaiban("ErrorID");
            }
            if (readCount == 0)
            {
                //同じファイルで何度目のエラーか取得
                readCount = GlobalMethod.getReadCount(errorUser);
            }
            ErrorMsg += "・" + msg;

            // エラーID、プログラム名、ファイル名、行番号、エラーメッセージ、エラーユーザー、エラー回数、種別？
            // 1レコードに複数のエラーを入れるようにする
            // row はシートを表すものとする
            // 1:調査概要
            // 2:担当部所
            // 3:調査品目一覧
            // 4:協力依頼書
            // 5:応援受付状況
            // 6:単品入力項目
            GlobalMethod.InsertErrorTable(errorId, "MadoguchiEXCEL", fileName, row, ErrorMsg, errorUser, readCount, 1);
        }

        // 履歴登録
        private void writeHistory(string historyMessage)
        {
            string pgmName = "Madoguchi";

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
                    ",'" + UserInfos[1] + "' " +
                    ",'" + UserInfos[2] + "' " +
                    ",'" + UserInfos[3] + "' " +
                    ",N'" + historyMessage + "' " +
                    ",'" + pgmName + methodName + "' ";

                // 採番した窓口IDが、エラーで存在しない場合があるので、IDの確認を行う
                string discript = "MadoguchiID ";
                string value = "MadoguchiID ";
                string table = "MadoguchiJouhou ";
                string where = " MadoguchiID = '" + saibanMadoguchiID + "' ";

                DataTable workDT = new DataTable();
                workDT = GlobalMethod.getData(discript, value, table, where);
                if (workDT != null && workDT.Rows.Count > 0)
                {
                    cmd.CommandText += "," + saibanMadoguchiID + " ";
                }
                else
                {
                    cmd.CommandText += ",NULL ";
                }
                cmd.CommandText += ",NULL " +
                    ",NULL " +
                    ",NULL " +
                    ",NULL " +
                    ",'" + tokuchoBangou + "' " +
                ")";

                cmd.ExecuteNonQuery();
                conn.Close();
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
            //    // height:628 → 1086・・・調査品目明細と合わせる
            //    // width:1864 → 3752
            //    btnGridSize.Text = "一覧縮小"; 
            //    c1FlexGrid1.Height = 1086;
            //    c1FlexGrid1.Width = 3752;
            //}
            //else
            //{
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid1.Height = 628;
            //    c1FlexGrid1.Width = 1864;
            //}
            gridSizeChange();
        }
        private void gridSizeChange()
        {
            string num = "";
            int bigHeight = 0;
            int bigWidth = 0;
            int smallHeight = 0;
            int smallWidth = 0;

            if (btnGridSize.Text == "一覧拡大")
            {
                num = GlobalMethod.GetCommonValue1("MADOGUCHI_GRID_BIG_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out bigHeight);
                    if (bigHeight == 0)
                    {
                        bigHeight = 1086;
                    }
                }
                num = GlobalMethod.GetCommonValue1("MADOGUCHI_GRID_BIG_WIDTH");
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
                //btnGridSize.Text = "一覧縮小";
                //c1FlexGrid1.Height = 1086;
                //c1FlexGrid1.Width = 3752;

                // height:628 → 1086・・・調査品目明細と合わせる
                // width:1864 → 3752
                btnGridSize.Text = "一覧縮小";
                c1FlexGrid1.Height = bigHeight;
                c1FlexGrid1.Width = bigWidth;

            }
            else
            {
                num = GlobalMethod.GetCommonValue1("MADOGUCHI_GRID_SMALL_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out smallHeight);
                    if (smallHeight == 0)
                    {
                        smallHeight = 628;
                    }
                }
                num = GlobalMethod.GetCommonValue1("MADOGUCHI_GRID_BIG_WIDTH");
                if (num != null)
                {
                    Int32.TryParse(num, out smallWidth);
                    if (smallWidth == 0)
                    {
                        smallWidth = 1864;
                    }
                }

                //btnGridSize.Text = "一覧拡大";
                //c1FlexGrid1.Height = 628;
                //c1FlexGrid1.Width = 1864;

                btnGridSize.Text = "一覧拡大";
                c1FlexGrid1.Height = smallHeight;
                c1FlexGrid1.Width = smallWidth;
            }
        }
    }
}
