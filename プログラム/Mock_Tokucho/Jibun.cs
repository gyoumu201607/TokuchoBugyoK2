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

namespace TokuchoBugyoK2
{
    public partial class Jibun : Form
    {
        private string pgmName = "Jibun";

        public string[] UserInfos;
        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();
        //private Boolean kensakukikanFlg = false;
        public Boolean ReSearch = false;
        public Boolean miteishutu = false;

        private Color errorBackColor = Color.FromArgb(255, 204, 255);

        public Jibun()
        {
            InitializeComponent();

            //マウスホイールイベント TODO
            // コンボボックスにマウスホイールイベントを付与
            this.item_Nendo.MouseWheel += item_MouseWheel;
            this.item_ChousaTantouBusho.MouseWheel += item_MouseWheel;
            this.item_MadoguchiBusho.MouseWheel += item_MouseWheel;
            this.item_FromTo.MouseWheel += item_MouseWheel;
            this.item_ShimekiriSentaku.MouseWheel += item_MouseWheel;
            this.item_Shintyokujyoukyo.MouseWheel += item_MouseWheel;
            this.item_Hyoujikensuu.MouseWheel += item_MouseWheel;
            this.item_Taisho.MouseWheel += item_MouseWheel;
            this.item_TantouJoukyo.MouseWheel += item_MouseWheel;
            //エントリ君修正STEP2
            this.ErrorMessage.Font = new System.Drawing.Font(this.ErrorMessage.Font.Name, float.Parse(GlobalMethod.GetCommonValue1("DSP_ERROR_FONTSIZE")));
        }

        private void Jibun_Load(object sender, EventArgs e)
        {
            //不具合No1355（1123）
            lblVersion.Text = GlobalMethod.GetCommonValue1("APL_VERSION");
            if (GlobalMethod.GetCommonValue1("BOOT_MODE") == "1")
            {
                lblBootMode.Text = GlobalMethod.GetCommonValue2("BOOT_MODE");
            }
            //ユーザ名を設定
            label3.Text = UserInfos[3] + "：" + UserInfos[1];

            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            //画像用Hashtable
            Hashtable imgMap = new Hashtable();
            Hashtable imgMap2 = new Hashtable();

            gridSizeChange();

            //ソート項目にアイコンを設定
            C1.Win.C1FlexGrid.CellRange cr;
            Bitmap bmp1 = new Bitmap("Resource/Image/SortIconDefalt.png");
            Bitmap bmpSort = new Bitmap(bmp1, bmp1.Width / 6, bmp1.Height / 6);
            cr = c1FlexGrid1.GetCellRange(0, 1);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;

            for (int i = 3; i < c1FlexGrid1.Cols.Count; i++)
            {
                if (i != 6)
                {
                    cr = c1FlexGrid1.GetCellRange(0, i);
                    cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
                    cr.Image = bmpSort;
                }
            }

            // 進捗アイコン
            imgMap2 = new Hashtable();
            imgMap2.Add("8", Image.FromFile("Resource/Image/shin_ao.png"));     // 報告済み
            imgMap2.Add("5", Image.FromFile("Resource/Image/greenT1.png"));     // 二次検証済み、または中止（二次検証済み）
            //imgMap2.Add("6", Image.FromFile("Resource/Image/greenT1.png"));     // 二次検証済み、または中止（中止）
            imgMap2.Add("6", Image.FromFile("Resource/Image/shin_ao.png"));     // 中止
            imgMap2.Add("7", Image.FromFile("Resource/Image/shin_midori.png")); // 担当者済み
            imgMap2.Add("1", Image.FromFile("Resource/Image/shin_dokuro.png")); // 締切日経過
            imgMap2.Add("2", Image.FromFile("Resource/Image/shin_aka.png"));    // 締切日が3日以内、かつ2次検証が完了していない
            imgMap2.Add("3", Image.FromFile("Resource/Image/shin_kiiro.png"));  // 締切日が1週間以内、かつ2次検証が完了していない
            imgMap2.Add("4", Image.FromFile("Resource/Image/blank2.png"));      // それ以外
            c1FlexGrid1.Cols[1].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[1].ImageMap = imgMap2;
            c1FlexGrid1.Cols[1].ImageAndText = false;

            //編集の画像切り替え
            imgMap = new Hashtable();
            imgMap.Add("0", Image.FromFile("Resource/Image/file_presentation1_g.png"));
            imgMap.Add("1", Image.FromFile("Resource/Image/file_presentation1.png"));
            c1FlexGrid1.Cols[2].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[2].ImageMap = imgMap;
            c1FlexGrid1.Cols[2].ImageAndText = false;

            //フォルダー画像アイコン
            imgMap = new Hashtable();
            imgMap.Add("0", Image.FromFile("Resource/Image/folder_gray_s.png"));
            imgMap.Add("1", Image.FromFile("Resource/Image/folder_yellow_s.png"));
            c1FlexGrid1.Cols[6].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[6].ImageMap = imgMap;
            c1FlexGrid1.Cols[6].ImageAndText = false;

            //不具合No1337（1094）対応　差戻分
            //グリッドの調査担当部所のキャプションに改行を入れる
            c1FlexGrid1.Cols[10].Caption = "調査担当" + Environment.NewLine + "部所";

            //コンボボックスセット
            set_combo();

            //検索条件初期化
            ClearForm();

            //一覧データ初期検索
            get_data(1);
            //kensakukikanFlg = false;
        }

        private void set_combo_shibu(string nendo)
        {
            //部所取得処理
            DataTable combodt;

            //窓口部所
            String discript = "Mst_Busho.BushokanriboKamei ";
            String value = "Mst_Busho.GyoumuBushoCD ";
            String table = "Mst_Busho";
            String where = "KashoShibuCD != '' AND GyoumuBushoCD != '999990' AND GyoumuBushoCD != '127900' AND BushoNewOld <= 1 AND BushoMadoguchiHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 ";
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
            }
            where += " ORDER BY BushoMadoguchiNarabijun";

            combodt = GlobalMethod.getData(discript, value, table, where);


            //コンボボックスデータ取得
            DataTable tmpdt = GlobalMethod.getData(discript, value, table, where);
            DataTable tmpdt2 = GlobalMethod.getData(discript, value, table, where);
            if (tmpdt != null)
            {
                //空白行追加
                DataRow dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            item_MadoguchiBusho.DataSource = tmpdt;
            item_MadoguchiBusho.DisplayMember = "Discript";
            item_MadoguchiBusho.ValueMember = "Value";


            //調査担当部所
            if (tmpdt2 != null)
            {
                //空白行追加
                DataRow dr = tmpdt2.NewRow();
                tmpdt2.Rows.InsertAt(dr, 0);
            }
            item_ChousaTantouBusho.DataSource = tmpdt2;
            item_ChousaTantouBusho.DisplayMember = "Discript";
            item_ChousaTantouBusho.ValueMember = "Value";
            //ユーザの部所セット
            item_ChousaTantouBusho.SelectedValue = UserInfos[2];
        }

        private void set_combo()
        {

            //コンボボックスの内容を設定

            DataTable combodt;
            System.Data.DataTable tmpdt;
            SortedList sl;

            //登録年度
            String discript = "NendoSeireki";
            String value = "NendoID ";
            String table = "Mst_Nendo";
            String where = "";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            item_Nendo.DataSource = combodt;
            item_Nendo.DisplayMember = "Discript";
            item_Nendo.ValueMember = "Value";

            //今年度を取得
            /*
            discript = "NendoSeireki";
            value = "NendoID";
            table = "Mst_Nendo";
            where = "Nendo_Sdate <= GETDATE() AND Nendo_EDate >= GETDATE()";
            //コンボボックスデータ取得
            DataTable dtYear = GlobalMethod.getData(discript, value, table, where);
            item_Nendo.SelectedValue = dtYear.Rows[0][0].ToString();
            */
            item_Nendo.SelectedValue = GlobalMethod.GetTodayNendo();

            //部所系コンボボックス取得
            //set_combo_shibu(dtYear.Rows[0][0].ToString());
            set_combo_shibu(GlobalMethod.GetTodayNendo());

            //検索期間の指定
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
                DataRow dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            item_FromTo.DataSource = tmpdt;
            item_FromTo.DisplayMember = "Discript";
            item_FromTo.ValueMember = "Value";

            //調査担当者のコンボボックス
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "主+副");
            tmpdt.Rows.Add(1, "主");
            tmpdt.Rows.Add(2, "副");
            item_Taisho.DataSource = tmpdt;
            item_Taisho.DisplayMember = "Discript";
            item_Taisho.ValueMember = "Value";

            //締切日の選択
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
                DataRow dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            item_ShimekiriSentaku.DataSource = tmpdt;
            item_ShimekiriSentaku.DisplayMember = "Discript";
            item_ShimekiriSentaku.ValueMember = "Value";

            //担当者状況
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            // VIPS　20220330　課題管理表No1294(982) ADD, CHANGE 担当者状況の追加・変更
            //tmpdt.Rows.Add(10, "依頼");
            tmpdt.Rows.Add(20, "調査開始");
            tmpdt.Rows.Add(30, "見積中");
            tmpdt.Rows.Add(40, "集計中");
            tmpdt.Rows.Add(50, "担当者済");
            tmpdt.Rows.Add(60, "一次検済");
            tmpdt.Rows.Add(70, "二次検済");
            //tmpdt.Rows.Add(68, "依頼・担当者済");
            //tmpdt.Rows.Add(69, "調査中・担当者済");
            tmpdt.Rows.Add(200, "調査中");
            tmpdt.Rows.Add(210, "調査中・担当者済");
            tmpdt.Rows.Add(300, "検証中");
            tmpdt.Rows.Add(310, "検証中・二次検済");
            tmpdt.Rows.Add(80, "中止");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            if (tmpdt != null)
            {
                DataRow dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            item_TantouJoukyo.DataSource = tmpdt;
            item_TantouJoukyo.DisplayMember = "Discript";
            item_TantouJoukyo.ValueMember = "Value";

            // 検索結果一覧グリッド内の担当者状況
            //担当者状況
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            //tmpdt.Rows.Add(0, "　");
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
            if (tmpdt != null)
            {
                DataRow dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }

            //該当グリッドのセルにセット
            //不具合No1337（1094）対応
            c1FlexGrid1.Cols[13].DataMap = sl;
            //c1FlexGrid1.Cols[14].DataMap = sl;

            //進捗状況
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(6, "超過");
            tmpdt.Rows.Add(5, "締め切りまで3日以内");
            tmpdt.Rows.Add(4, "締め切りまで1週間以内");
            tmpdt.Rows.Add(3, "締め切りまで1週間以上 または 中止");
            tmpdt.Rows.Add(2, "担当者済");
            tmpdt.Rows.Add(7, "一次検済");
            tmpdt.Rows.Add(1, "二次検済");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            if (tmpdt != null)
            {
                DataRow dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            item_Shintyokujyoukyo.DataSource = tmpdt;
            item_Shintyokujyoukyo.DisplayMember = "Discript";
            item_Shintyokujyoukyo.ValueMember = "Value";

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
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[15].DataMap = sl;

            // 調査区分
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "無");
            tmpdt.Rows.Add(1, "有");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //グリッドにセット
            //c1FlexGrid1.Cols[16].DataMap = sl; // 調査区分 自部所
            //c1FlexGrid1.Cols[17].DataMap = sl; // 調査区分 支→支
            //c1FlexGrid1.Cols[18].DataMap = sl; // 調査区分 本→支
            //c1FlexGrid1.Cols[19].DataMap = sl; // 調査区分 支→本

            // 管理帳票印刷コンボボックス
            discript = "PrintName";
            value = "PrintListID";
            table = "Mst_PrintList";
            //where = "";
            where = "MENU_ID = 400 AND PrintBunruiCD = 1 AND PrintDelFlg <> 1 ORDER BY PrintListNarabijun";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            //if (combodt != null)
            //{
            //    DataRow dr = combodt.NewRow();
            //    combodt.Rows.InsertAt(dr, 0);
            //}
            item_tyouhyouInsatu.DataSource = combodt;
            item_tyouhyouInsatu.DisplayMember = "Discript";
            item_tyouhyouInsatu.ValueMember = "Value";

        }

        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            //編集ボタン押下時の処理
            var hti = this.c1FlexGrid1.HitTest(new Point(e.X, e.Y));
            //if (hti.Column == 2)
            if (hti.Column == 2 && hti.Row > 0)
            {
                ReSearch = true;
                Jibun_Input form = new Jibun_Input();
                form.MadoguchiID = c1FlexGrid1[hti.Row, 29].ToString();
                form.ChousaCD = c1FlexGrid1[hti.Row, 30].ToString();
                form.UserInfos = UserInfos;
                form.Show(this);
                this.Hide();
            }
            // フォルダアイコン
            if (hti.Column == 6 & hti.Row >= 1)
            {
                if (c1FlexGrid1[hti.Row, 6].ToString() == "1")
                {
                    // フォルダ表示
                    System.Diagnostics.Process.Start("EXPLORER.EXE", GlobalMethod.GetPathValid(c1FlexGrid1[hti.Row, 31].ToString()));
                }
            }
        }

        //検索時必須チェック
        private Boolean search_required()
        {
            Boolean requiredFlg = true;

            //背景を白に戻す
            item_Nendo.BackColor = Color.FromArgb(255, 255, 255);

            //登録年度が空だった場合
            if (String.IsNullOrEmpty(item_Nendo.Text))
            {
                requiredFlg = false;
                item_Nendo.BackColor = Color.FromArgb(255, 204, 255);
            }

            //requiredFlgがfalseの場合
            if (!requiredFlg)
            {
                //「必須入力項目が入力されていません。」
                set_error(GlobalMethod.GetMessage("E20901", ""));
            }

            return requiredFlg;
        }

        private void get_data(int page)
        {

            ListData = new DataTable();

            string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            //データ取得処理
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();

                //SQL生成
                cmd.CommandText = "SELECT " +
                  "mj.MadoguchiShinchokuJoukyou " + //進捗状況　0
                  ",mj.MadoguchiHoukokuzumi " + //報告済み
                  ",mj.MadoguchiKanryou " + //完了 2
                  ",mj.MadoguchiUketsukeBangou + '-' + MadoguchiUketsukeBangouEdaban as tokuchoBan　" + //特調番号（枝番付き）3
                  ",mj.MadoguchiHachuuKikanmei " + //発注者名・課名
                  ",mj.MadoguchiShiryouHolder " + // 5
                  ",mj.MadoguchiTourokubi " + //登録日 6
                  ",mjmc.MadoguchiL1ChousaShimekiribi " + //締切日
                  ",mjmc.MadoguchiL1ChousaBushoCD " + //調査担当部所CD　非表示 9
                  ",mb2.BushokanriboKamei " + //調査担当部所
                  ",mjmc.MadoguchiL1ChousaTantoushaCD " + //調査担当者CD　非表示
                  ",mjmc.MadoguchiL1ChousaTantousha " + //調査担当者
                  ",mjmc.MadoguchiL1ChousaShinchoku " + //担当者状況
                  ",mjmc.MadoguchiL1Memo " + //メモ 24
                  ",CASE mj.MadoguchiChousaShubetsu WHEN 1 THEN '単品' WHEN 2 THEN '一般' WHEN 3 THEN '単契' ELSE ' ' END AS MadoguchiChousaShubetsu" + //調査種別 14 //非表示列
                  ",CASE mj.MadoguchiChousaKubunJibusho WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunJibusho" + //調査区分　自部所 //非表示列
                  ",CASE mj.MadoguchiChousaKubunShibuShibu WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunShibuShibu" + //調査区分　支部→支部 //非表示列
                  ",CASE mj.MadoguchiChousaKubunHonbuShibu WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunHonbuShibu" + //調査区分　本部→支部 //非表示列
                  ",CASE mj.MadoguchiChousaKubunShibuHonbu WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunShibuHonbu" + //調査区分　支部→本部 //非表示列
                  ",mj.MadoguchiKanriBangou " + //管理番号 19
                  ",mj.MadoguchiBushoRenban " + //部所連番
                  ",mj.MadoguchiGyoumuMeishou " + //業務名称
                  ",mj.MadoguchiKoujiKenmei " + //工事件名
                  ",sb.ShibuBikou "+
                  ",mj.MadoguchiChousaHinmoku " + //調査品目
                  ",mb.BushokanriboKamei " + //窓口部所
                  ",mc.ChousainMei " + //窓口担当者
                  ",mj.MadoguchiOuenUketsukebi " + //応援受付日
                  ",mj.MadoguchiID " +
                  ",mjmc.MadoguchiL1ChousaCD " +
                  ",mj.MadoguchiShiryouHolder " + // 30
                  ",mj.MadoguchiHoukokuzumi " + // 31:報告済み

                // 32:進捗アイコンの判定用
                ", " +
                "CASE " +
                "WHEN mj.MadoguchiHoukokuzumi = 1 THEN 8 " +
                "WHEN mj.MadoguchiHoukokuzumi != 1 THEN " +
                "     CASE " +
                "         WHEN mjmc.MadoguchiL1ChousaShinchoku = 80 THEN 6 " +
                "         WHEN mjmc.MadoguchiL1ChousaShinchoku = 70 THEN 5 " +
                "         WHEN mjmc.MadoguchiL1ChousaShinchoku = 50 THEN 7 " +
                "         WHEN mjmc.MadoguchiL1ChousaShinchoku = 60 THEN 7 " + // 一次検済
                "     ELSE " +
                "         CASE " +
                "              WHEN mjmc.MadoguchiL1ChousaShimekiribi < '" + DateTime.Today + "' THEN 1 " +
                "              WHEN mjmc.MadoguchiL1ChousaShimekiribi <= '" + DateTime.Today.AddDays(3) + "' THEN 2 " +
                "              WHEN mjmc.MadoguchiL1ChousaShimekiribi <= '" + DateTime.Today.AddDays(7) + "' THEN 3 " +
                "         ELSE 4 " +
                "         END " +
                "     END " +
                "END as Shinchoku " +
                "FROM MadoguchiJouhouMadoguchiL1Chou mjmc " +
                "INNER JOIN MadoguchiJouhou mj " +
                "  ON mj.MadoguchiID = mjmc.MadoguchiID " +
                " AND mjmc.MadoguchiL1DeleteFlag = 0 " +
                "LEFT JOIN ShibuBikou sb " +
                "  ON mj.MadoguchiID = sb.MadoguchiID " +
                " AND sb.ShibuBikouBushoKanriboBushoCD = mjmc.MadoguchiL1ChousaBushoCD " +
                " AND sb.ShinDeleteFlag = 0 " +
                "LEFT JOIN Mst_Chousain mc ON  mj.MadoguchiTantoushaCD = mc.KojinCD " +
                "LEFT JOIN Mst_Busho mb2 ON mjmc.MadoguchiL1ChousaBushoCD = mb2.GyoumuBushoCD " +
                "LEFT JOIN Mst_Busho mb ON mj.MadoguchiTantoushaBushoCD = mb.GyoumuBushoCD " +
                "WHERE mj.MadoguchiDeleteFlag = 0 " +
                " AND mj.MadoguchiSystemRenban > 0 ";

                ////SQL生成
                //cmd.CommandText = "SELECT " +
                //  "mj.MadoguchiShinchokuJoukyou " + //進捗状況　0
                //  ",mj.MadoguchiHoukokuzumi " + //報告済み
                //  ",mj.MadoguchiKanryou " + //完了 2
                //  ",mj.MadoguchiUketsukeBangou + '-' + MadoguchiUketsukeBangouEdaban as tokuchoBan　" + //特調番号（枝番付き）3
                //  ",mj.MadoguchiHachuuKikanmei " + //発注者名・課名
                //  ",mj.MadoguchiShiryouHolder " + // 5
                //  ",mj.MadoguchiTourokubi " + //登録日 6
                //  ",mj.MadoguchiOuenUketsukebi " + //応援受付日
                //                                   //",mj.MadoguchiShimekiribi " + //締切日
                //  ",mjmc.MadoguchiL1ChousaShimekiribi " + //締切日
                //  ",mjmc.MadoguchiL1ChousaBushoCD " + //調査担当部所CD　非表示 9
                //                                      //",mjmc.MadoguchiL1ChousaBusho " + //調査担当部所
                //  ",mb2.BushokanriboKamei " + //調査担当部所
                //  ",mjmc.MadoguchiL1ChousaTantoushaCD " + //調査担当者CD　非表示
                //  ",mjmc.MadoguchiL1ChousaTantousha " + //調査担当者
                //  ",mjmc.MadoguchiL1ChousaShinchoku " + //担当者状況
                //  ",CASE mj.MadoguchiChousaShubetsu WHEN 1 THEN '単品' WHEN 2 THEN '一般' WHEN 3 THEN '単契' ELSE ' ' END AS MadoguchiChousaShubetsu" + //調査種別 14
                //  ",CASE mj.MadoguchiChousaKubunJibusho WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunJibusho" + //調査区分　自部所
                //  ",CASE mj.MadoguchiChousaKubunShibuShibu WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunShibuShibu" + //調査区分　支部→支部
                //  ",CASE mj.MadoguchiChousaKubunHonbuShibu WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunHonbuShibu" + //調査区分　本部→支部
                //  ",CASE mj.MadoguchiChousaKubunShibuHonbu WHEN 1 THEN '有' ELSE '無' END AS MadoguchiChousaKubunShibuHonbu" + //調査区分　支部→本部
                //  ",mj.MadoguchiKanriBangou " + //管理番号 19
                //  ",mj.MadoguchiBushoRenban " + //部所連番
                //  ",mj.MadoguchiGyoumuMeishou " + //業務名称
                //  ",mj.MadoguchiKoujiKenmei " + //工事件名
                //  ",sb.ShibuBikou ";  //部所備考
                //                      //",(select TOP 1 ShibuBikou from ShibuBikou where ShibuBikouBushoKanriboBushoCD = " + UserInfos[2] + " and MadoguchiID = mj.MadoguchiID and ShinDeleteFlag = 0 ";
                //                      //  ",(select TOP 1 ShibuBikou from ShibuBikou where ShibuBikouBushoKanriboBushoCD = mjmc.MadoguchiL1ChousaBushoCD and MadoguchiID = mj.MadoguchiID and ShinDeleteFlag = 0 ";
                //                      ////部所備考
                //                      //if (item_BushoBikou.Text != "")
                //                      //{
                //                      //    cmd.CommandText += "AND ShibuBikou LIKE '%" + GlobalMethod.ChangeSqlText(item_BushoBikou.Text, 1) + "%' ESCAPE '\\' ";
                //                      //}
                //                      //cmd.CommandText += ") ";

                //cmd.CommandText += ",mjmc.MadoguchiL1Memo " + //メモ 24
                //  ",mj.MadoguchiChousaHinmoku " + //調査品目
                //  ",mc.ChousainMei " + //窓口担当者
                //  ",mb.BushokanriboKamei " + //窓口部所
                //  ",mj.MadoguchiID " +
                //  ",mjmc.MadoguchiL1ChousaCD " +
                //  ",mj.MadoguchiShiryouHolder " + // 30
                //  ",mj.MadoguchiHoukokuzumi " + // 31:報告済み

                //// 32:進捗アイコンの判定用
                //", " +
                //"CASE " +
                //"WHEN mj.MadoguchiHoukokuzumi = 1 THEN 8 " +
                //"WHEN mj.MadoguchiHoukokuzumi != 1 THEN " +
                //"     CASE " +
                //"         WHEN mjmc.MadoguchiL1ChousaShinchoku = 80 THEN 6 " +
                //"         WHEN mjmc.MadoguchiL1ChousaShinchoku = 70 THEN 5 " +
                //"         WHEN mjmc.MadoguchiL1ChousaShinchoku = 50 THEN 7 " +
                //"         WHEN mjmc.MadoguchiL1ChousaShinchoku = 60 THEN 7 " + // 一次検済
                //"     ELSE " +
                //"         CASE " +
                //"              WHEN mjmc.MadoguchiL1ChousaShimekiribi < '" + DateTime.Today + "' THEN 1 " +
                //"              WHEN mjmc.MadoguchiL1ChousaShimekiribi <= '" + DateTime.Today.AddDays(3) + "' THEN 2 " +
                //"              WHEN mjmc.MadoguchiL1ChousaShimekiribi <= '" + DateTime.Today.AddDays(7) + "' THEN 3 " +
                //"         ELSE 4 " +
                //"         END " +
                //"     END " +
                //"END as Shinchoku " +
                //"FROM MadoguchiJouhouMadoguchiL1Chou mjmc " +
                //"INNER JOIN MadoguchiJouhou mj " +
                //"  ON mj.MadoguchiID = mjmc.MadoguchiID " +
                //" AND mjmc.MadoguchiL1DeleteFlag = 0 " +
                //"LEFT JOIN ShibuBikou sb " +
                //"  ON mj.MadoguchiID = sb.MadoguchiID " +
                //" AND sb.ShibuBikouBushoKanriboBushoCD = mjmc.MadoguchiL1ChousaBushoCD " +
                //" AND sb.ShinDeleteFlag = 0 " +
                //"LEFT JOIN Mst_Chousain mc ON  mj.MadoguchiTantoushaCD = mc.KojinCD " +
                //"LEFT JOIN Mst_Busho mb2 ON mjmc.MadoguchiL1ChousaBushoCD = mb2.GyoumuBushoCD " +
                //"LEFT JOIN Mst_Busho mb ON mj.MadoguchiTantoushaBushoCD = mb.GyoumuBushoCD " +
                ////"WHERE MadoguchiTourokuNendo = '" + item_Nendo.SelectedValue.ToString() + "' " +
                ////  "AND mj.MadoguchiDeleteFlag = 0 ";
                //"WHERE mj.MadoguchiDeleteFlag = 0 " +
                //" AND mj.MadoguchiSystemRenban > 0 ";

                // 当年度
                if (item_NendoOptionTounen.Checked)
                {
                    cmd.CommandText += "AND MadoguchiTourokuNendo = '" + item_Nendo.SelectedValue.ToString() + "' ";
                }

                // 3年以内
                if (item_NendoOption3Nen.Checked)
                {
                    int.TryParse(item_Nendo.SelectedValue.ToString(), out int w_Nendo);
                    cmd.CommandText += "AND (MadoguchiTourokuNendo = '" + w_Nendo.ToString() + "' ";
                    w_Nendo = w_Nendo - 1;
                    cmd.CommandText += "OR MadoguchiTourokuNendo = '" + w_Nendo.ToString() + "' ";
                    w_Nendo = w_Nendo - 1;
                    cmd.CommandText += "OR MadoguchiTourokuNendo = '" + w_Nendo.ToString() + "') ";
                }

                //不具合No1337（1094）対応
                //if (item_ChousaKbnJibusho.Checked || item_ChousaKbnShibuShibu.Checked || item_ChousaKbnHonbuShibu.Checked || item_ChousaKbnShibuHonbu.Checked)
                //{
                //    // OR追加フラグ true:OR追加
                //    //Boolean OrAddFlg = false;

                //    //cmd.CommandText += "AND (";
                //    ////調査区分　自部所
                //    //if (item_ChousaKbnJibusho.Checked)
                //    //{
                //    //    cmd.CommandText += "mj.MadoguchiChousaKubunJibusho = 1 ";
                //    //    OrAddFlg = true;
                //    //}

                //    ////調査区分　支部→支部
                //    //if (item_ChousaKbnShibuShibu.Checked)
                //    //{
                //    //    if (OrAddFlg)
                //    //    {
                //    //        //cmd.CommandText += "OR ";
                //    //        cmd.CommandText += "AND ";
                //    //    }
                //    //    cmd.CommandText += "mj.MadoguchiChousaKubunShibuShibu = 1 ";
                //    //    OrAddFlg = true;
                //    //}

                //    ////調査区分　本部→支部
                //    //if (item_ChousaKbnHonbuShibu.Checked)
                //    //{
                //    //    if (OrAddFlg)
                //    //    {
                //    //        //cmd.CommandText += "OR ";
                //    //        cmd.CommandText += "AND ";
                //    //    }
                //    //    cmd.CommandText += "mj.MadoguchiChousaKubunHonbuShibu = 1 ";
                //    //    OrAddFlg = true;
                //    //}

                //    ////調査区分　支部→本部
                //    //if (item_ChousaKbnShibuHonbu.Checked)
                //    //{
                //    //    if (OrAddFlg)
                //    //    {
                //    //        //cmd.CommandText += "OR ";
                //    //        cmd.CommandText += "AND ";
                //    //    }
                //    //    cmd.CommandText += "mj.MadoguchiChousaKubunShibuHonbu = 1 ";
                //    //    OrAddFlg = true;
                //    //}
                //    //cmd.CommandText += ")";
                //    cmd.CommandText += "AND (";
                //    //調査区分　自部所
                //    if (item_ChousaKbnJibusho.Checked)
                //    {
                //        cmd.CommandText += " MadoguchiChousaKubunJibusho = 1 ";
                //    }
                //    else
                //    {
                //        cmd.CommandText += " MadoguchiChousaKubunJibusho = 0 ";
                //    }

                //    cmd.CommandText += "AND ";
                //    //調査区分　支部→支部
                //    if (item_ChousaKbnShibuShibu.Checked)
                //    {
                //        cmd.CommandText += " MadoguchiChousaKubunShibuShibu = 1 ";
                //    }
                //    else
                //    {
                //        cmd.CommandText += " MadoguchiChousaKubunShibuShibu = 0 ";
                //    }
                //    cmd.CommandText += "AND ";
                //    //調査区分　本部→支部
                //    if (item_ChousaKbnHonbuShibu.Checked)
                //    {
                //        cmd.CommandText += " MadoguchiChousaKubunHonbuShibu = 1 ";
                //    }
                //    else
                //    {
                //        cmd.CommandText += " MadoguchiChousaKubunHonbuShibu = 0 ";
                //    }

                //    cmd.CommandText += "AND ";
                //    //調査区分　支部→本部
                //    if (item_ChousaKbnShibuHonbu.Checked)
                //    {
                //        cmd.CommandText += " MadoguchiChousaKubunShibuHonbu = 1 ";
                //    }
                //    else
                //    {
                //        cmd.CommandText += " MadoguchiChousaKubunShibuHonbu = 0 ";
                //    }
                //    cmd.CommandText += ")";
                //}

                //// 調査区分 自部所
                //if (item_ChousaKbnJibusho.Checked)
                //{
                //    cmd.CommandText += "AND mj.MadoguchiChousaKubunJibusho = 1 ";
                //}
                //// 調査区分 支→支
                //if (item_ChousaKbnShibuShibu.Checked)
                //{
                //    cmd.CommandText += "AND mj.MadoguchiChousaKubunShibuShibu = 1 ";
                //}
                //// 調査区分 本→支
                //if (item_ChousaKbnHonbuShibu.Checked)
                //{
                //    cmd.CommandText += "AND mj.MadoguchiChousaKubunHonbuShibu = 1 ";
                //}
                //// 調査区分 支→本
                //if (item_ChousaKbnShibuHonbu.Checked)
                //{
                //    cmd.CommandText += "AND mj.MadoguchiChousaKubunShibuHonbu = 1 ";
                //}

                //発注者名・課名
                if (item_HachushaKamei.Text != "")
                {
                    cmd.CommandText += "AND mj.MadoguchiHachuuKikanmei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_HachushaKamei.Text, 1) + "%' ESCAPE '\\' ";
                }

                //特調番号
                if (item_TokuchoNo.Text != "")
                {
                    cmd.CommandText += "AND CONCAT(mj.MadoguchiUketsukeBangou,'-', mj.MadoguchiUketsukeBangouEdaban) COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_TokuchoNo.Text, 1) + "%' ESCAPE '\\' ";
                }

                //調査担当部所
                if (!String.IsNullOrEmpty(item_ChousaTantouBusho.Text))
                {
                    cmd.CommandText += "AND mjmc.MadoguchiL1ChousaBushoCD = '" + item_ChousaTantouBusho.SelectedValue.ToString() + "' ";
                }

                //窓口部所
                if (!String.IsNullOrEmpty(item_MadoguchiBusho.Text))
                {
                    cmd.CommandText += "AND mj.MadoguchiTantoushaBushoCD = '" + item_MadoguchiBusho.SelectedValue.ToString() + "' ";
                }

                //業務名称
                //不具合No1337（1094）対応
                //if (item_Gyoumumei.Text != "")
                //{
                //    cmd.CommandText += "AND mj.MadoguchiGyoumuMeishou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_Gyoumumei.Text, 1) + "%' ESCAPE '\\' ";
                //}

                //締切日 From
                if (item_DateFrom.CustomFormat == "")
                {
                    //cmd.CommandText += "AND mj.MadoguchiShimekiribi >= '" + Get_DateTimePicker("item_DateFrom") + "' ";
                    cmd.CommandText += "AND mjmc.MadoguchiL1ChousaShimekiribi >= '" + Get_DateTimePicker("item_DateFrom") + "' ";
                }
                //締切日 To
                if (item_DateTo.CustomFormat == "")
                {
                    //1日加算
                    //string dateStr = Get_DateTimePicker("item_DateTo");
                    //DateTime dateTime = DateTime.Parse(dateStr);
                    //dateTime = dateTime.AddDays(1);
                    //cmd.CommandText += "AND mj.MadoguchiShimekiribi < '" + dateTime + "' ";
                    //cmd.CommandText += "AND mjmc.MadoguchiL1ChousaShimekiribi < '" + dateTime + "' ";
                    cmd.CommandText += "AND mjmc.MadoguchiL1ChousaShimekiribi <= '" + Get_DateTimePicker("item_DateTo") + "' ";
                }

                //管理番号
                //不具合No1337（1094）対応
                //if (item_KanriBangou.Text != "")
                //{
                //    cmd.CommandText += "AND mj.MadoguchiKanriBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_KanriBangou.Text, 1) + "%' ESCAPE '\\' ";
                //}

                //調査担当者 名前の文字列検索
                if (item_ChousaTantousha.Text != "")
                {
                    cmd.CommandText += "AND mjmc.MadoguchiL1ChousaTantousha COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_ChousaTantousha.Text, 1) + "%' ESCAPE '\\' ";
                }

                // 対象が主+副のとき
                // 全件対象

                //対象が主のとき
                if ("1".Equals(item_Taisho.SelectedValue.ToString()))
                {
                    cmd.CommandText += "AND mjmc.MadoguchiL1ShuTantouFlag = 1 ";
                    //cmd.CommandText += "AND mjmc.MadoguchiL1FukuTantouFlag != 1 ";
                }

                //対象が副のとき
                //if ("2".Equals(item_Taisho.SelectedValue.ToString()))
                if ("2".Equals(item_Taisho.SelectedValue.ToString()))
                {
                    //cmd.CommandText += "AND mjmc.MadoguchiL1ShuTantouFlag != 1 ";
                    cmd.CommandText += "AND mjmc.MadoguchiL1FukuTantouFlag = 1 ";
                }

                //本部単品
                //不具合No1337（1094）対応
                //if (item_HonbuTanpin.Checked)
                //{
                //    cmd.CommandText += "AND mj.MadoguchiHonbuTanpinflg = 1 ";
                //}

                //工事件名
                if (item_Koujikenmei.Text != "")
                {
                    cmd.CommandText += "AND mj.MadoguchiKoujiKenmei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_Koujikenmei.Text, 1) + "%' ESCAPE '\\' ";
                }

                //メモ
                if (item_Memo.Text != "")
                {
                    cmd.CommandText += "AND mjmc.MadoguchiL1Memo COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_Memo.Text, 1) + "%' ESCAPE '\\' ";
                }

                //担当者状況
                if (miteishutu)
                {
                    //未提出フラグtrueのとき　依頼　状態のものは出さない
                    cmd.CommandText += "AND mjmc.MadoguchiL1ChousaShinchoku <> 10 ";
                }
                if (!String.IsNullOrEmpty(item_TantouJoukyo.Text))
                {
                    //MessageBox.Show("");
                    int tantoushaJoukyo = int.Parse(item_TantouJoukyo.SelectedValue.ToString());
                    //担当者状況が依頼中、集計中、担当者済、二次検済、中止の場合
                    //if (!String.IsNullOrEmpty(item_TantouJoukyo.Text) && (tantoushaJoukyo < 60 || tantoushaJoukyo <= 70))
                    if (tantoushaJoukyo <= 60 || tantoushaJoukyo == 70 || tantoushaJoukyo == 80)
                    {
                        cmd.CommandText += "AND mjmc.MadoguchiL1ChousaShinchoku = " + tantoushaJoukyo + " ";
                    }

                    // VIPS　20220330　課題管理表No1294(982) DEL 依頼・担当者済の項目削除
                    //担当者状況が依頼・担当者済の場合
                    //if (!String.IsNullOrEmpty(item_TantouJoukyo.Text) && tantoushaJoukyo == 68)
                    //if (tantoushaJoukyo == 68)
                    //{
                    //    cmd.CommandText += "AND (mjmc.MadoguchiL1ChousaShinchoku = 10 " +
                    //        "OR mjmc.MadoguchiL1ChousaShinchoku = 50 )";
                    //}

                    // VIPS　20220330　課題管理表No1294(982) ADD 調査中の追加
                    //担当者状況が調査中の場合
                    if (!String.IsNullOrEmpty(item_TantouJoukyo.Text) && tantoushaJoukyo == 200)
                    {
                        cmd.CommandText += "AND (mjmc.MadoguchiL1ChousaShinchoku = 20 " +
                            "OR mjmc.MadoguchiL1ChousaShinchoku = 30" +
                            "OR mjmc.MadoguchiL1ChousaShinchoku = 40 )";
                    }

                    // VIPS　20220330　課題管理表No1294(982) CHANGE 調査中・担当者済の条件変更
                    //担当者状況が調査中・担当者済の場合
                    if (!String.IsNullOrEmpty(item_TantouJoukyo.Text) && tantoushaJoukyo == 210)
                    {
                        cmd.CommandText += "AND (mjmc.MadoguchiL1ChousaShinchoku = 20 " +
                            "OR mjmc.MadoguchiL1ChousaShinchoku = 30" +
                            "OR mjmc.MadoguchiL1ChousaShinchoku = 40" +
                            "OR mjmc.MadoguchiL1ChousaShinchoku = 50 )";
                    }

                    // VIPS　20220330　課題管理表No1294(982) ADD 検証中の追加
                    //担当者状況が検証中の場合
                    if (!String.IsNullOrEmpty(item_TantouJoukyo.Text) && tantoushaJoukyo == 300)
                    {
                        cmd.CommandText += "AND (mjmc.MadoguchiL1ChousaShinchoku = 50 " +
                            "OR mjmc.MadoguchiL1ChousaShinchoku = 60 )";
                    }

                    // VIPS　20220330　課題管理表No1294(982) ADD 検証中・二次検済の追加
                    //担当者状況が検証中・二次検済の場合
                    if (!String.IsNullOrEmpty(item_TantouJoukyo.Text) && tantoushaJoukyo == 310)
                    {
                        cmd.CommandText += "AND (mjmc.MadoguchiL1ChousaShinchoku = 50 " +
                            "OR mjmc.MadoguchiL1ChousaShinchoku = 60" +
                            "OR mjmc.MadoguchiL1ChousaShinchoku = 70 )";
                    }

                    ////担当者状況が一次検済の場合
                    //if (!String.IsNullOrEmpty(item_TantouJoukyo.Text) && tantoushaJoukyo == 60)
                    //{
                    //    cmd.CommandText += "AND mjmc.MadoguchiL1ChousaShinchoku = 60 ";
                    //}
                }
                //部所備考
                //不具合No1337（1094）対応
                //if (item_BushoBikou.Text != "")
                //{
                //    cmd.CommandText += "AND sb.ShibuBikou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_BushoBikou.Text, 1) + "%' ESCAPE '\\' ";
                //}

                //進捗Value
                String w_jyokyou = "";
                //締切日計算用
                //String workdayFrom = "0";
                //String workdayTo = "0";
                //DateTime w_Simekiribi6 = DateTime.Today;

                DateTime ShimekiriFrom = DateTime.MinValue;
                DateTime ShimekiriTo = DateTime.MinValue;
                DateTime w_Simekiribi7 = DateTime.MinValue;

                //進捗状況
                if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() != "")
                {
                    switch (item_Shintyokujyoukyo.SelectedValue.ToString())
                    {
                        // 2次検済み
                        case "1":
                            w_jyokyou = "70";
                            break;
                        // 担当者済
                        case "2":
                            w_jyokyou = "50";
                            break;
                        // 締切日が1週間をこえる もしくは中止
                        case "3":
                            w_Simekiribi7 = DateTime.Today.AddDays(7);
                            break;
                        // 締切日が1週間以内
                        case "4":
                            //workdayFrom = "-1";
                            //workdayTo = "7";
                            ShimekiriFrom = DateTime.Today;
                            ShimekiriTo = ShimekiriFrom.AddDays(7);
                            break;
                        // 締切日が3日以内
                        case "5":
                            //workdayFrom = "-1";
                            //workdayTo = "3";
                            ShimekiriFrom = DateTime.Today;
                            ShimekiriTo = ShimekiriFrom.AddDays(3);
                            break;
                        // 締切日が超過
                        case "6":
                            //workdayTo = "-1";
                            ShimekiriTo = DateTime.Today.AddDays(-1);
                            break;
                        // 1次検済み
                        case "7":
                            w_jyokyou = "60";
                            break;
                        default:
                            break;
                    }
                }

                // 締切日付計算
                //DateTime w_SimekiribiFrom = DateTime.Today.AddDays(int.Parse(workdayFrom));
                //DateTime w_SimekiribiTo = DateTime.Today.AddDays(int.Parse(workdayTo));
                DateTime w_SimekiribiFrom = ShimekiriFrom.Date;
                DateTime w_SimekiribiTo = ShimekiriTo.Date;

                // 進捗状況が一次検済か二次検済か担当者済の場合
                if (w_jyokyou != "" && (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "1" || item_Shintyokujyoukyo.SelectedValue.ToString() == "2" || item_Shintyokujyoukyo.SelectedValue.ToString() == "7"))
                {
                    //cmd.CommandText += "and mj.MadoguchiShinchokuJoukyou = " + w_jyokyou + " ";
                    cmd.CommandText += "and mjmc.MadoguchiL1ChousaShinchoku = " + w_jyokyou + " ";
                }

                //進捗状況が「締め切りまで1週間以上 または 中止（完了でない）」
                if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "3")
                {
                    //cmd.CommandText += "and (mj.MadoguchiShimekiribi >= '" + w_Simekiribi6 + "') or(mj.MadoguchiShinchokuJoukyou = 80) ";
                    //cmd.CommandText += "and mj.MadoguchiShinchokuJoukyou < 50	or mj.MadoguchiShinchokuJoukyou = 80 ";
                    //cmd.CommandText += "and mj.MadoguchiKanryou <> 1 ";
                    cmd.CommandText += "and ((mjmc.MadoguchiL1ChousaShinchoku < 50 and mjmc.MadoguchiL1ChousaShimekiribi >= '" + w_Simekiribi7 + "') ";
                    cmd.CommandText += "or mjmc.MadoguchiL1ChousaShinchoku = 80) ";

                }
                //進捗状況が 締切日が1週間以内、3日以内、超過のとき
                //if (workdayFrom != "0")
                if (w_SimekiribiFrom != DateTime.MinValue.Date)
                {
                    //cmd.CommandText += "and MadoguchiShimekiribi > '" + w_SimekiribiFrom + "' ";
                    //cmd.CommandText += "and mjmc.MadoguchiL1ChousaShimekiribi > '" + w_SimekiribiFrom + "' ";
                    cmd.CommandText += "and mjmc.MadoguchiL1ChousaShimekiribi >= '" + w_SimekiribiFrom + "' ";
                }
                //if (workdayTo != "0")
                if (w_SimekiribiTo != DateTime.MinValue.Date)
                {
                    //cmd.CommandText += "and MadoguchiShimekiribi <= '" + w_SimekiribiTo + "' ";
                    cmd.CommandText += "and mjmc.MadoguchiL1ChousaShimekiribi <= '" + w_SimekiribiTo + "' ";
                }

                // 窓口完了が立っていないのが条件
                // 2次検証済ではない　のと　完了 中止ではないのが条件
                if (item_Shintyokujyoukyo.SelectedValue != null && (item_Shintyokujyoukyo.SelectedValue.ToString() == "4" || item_Shintyokujyoukyo.SelectedValue.ToString() == "5" || item_Shintyokujyoukyo.SelectedValue.ToString() == "6"))
                {
                    //cmd.CommandText += "and MadoguchiKanryou <> 1 ";
                    //cmd.CommandText += "and MadoguchiShinchokuJoukyou < 50 ";
                    cmd.CommandText += "and mjmc.MadoguchiL1ChousaShinchoku < 50 ";
                }

                // 1206 超過の場合は、報告済みは除外する
                if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "6")
                {
                    cmd.CommandText += "and mj.MadoguchiHoukokuzumi = 0 ";
                }

                //ORDER BY ページング
                cmd.CommandText += "ORDER BY mjmc.MadoguchiL1ChousaShimekiribi DESC,mj.MadoguchiKanriBangou , tokuchoBan ";

                //スキップするレコードの数
                int skipCount = 0;
                //1ページより後ろならスキップするレコードを計算
                if (page > 1)
                {
                    skipCount = (page - 1) * int.Parse(item_Hyoujikensuu.Text);
                }

                //cmd.CommandText += "OFFSET " + skipCount + " ROWS " +
                //    "FETCH NEXT " + item_Hyoujikensuu.Text + " ROWS ONLY; ";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                ListData.Clear();
                sda.Fill(ListData);

                // 0件の場合
                if (ListData.Rows.Count == 0)
                {
                    set_error("", 0);

                    //1ページ目でデータがないとき
                    if (page == 1)
                    {
                        // I20001:該当データは0件です。
                        set_error(GlobalMethod.GetMessage("I20001", ""));
                    }
                }

                Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / int.Parse(item_Hyoujikensuu.Text))).ToString();
                Paging_now.Text = (page).ToString();
                //データセット
                set_data(page);
                //ページングボタン制御
                set_page_enabled(page, int.Parse(Paging_all.Text));

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

        }

        private void set_data(int pagenum)
        {
            //Gridデータセット処理
            c1FlexGrid1.Rows.Count = 1;
            c1FlexGrid1.AllowAddNew = true;

            if (ListData.Rows.Count > 0)
            {

                //c1FlexGrid1.Cols[6].Style.BackgroundImage = Image.FromFile("Resource/Image/folder_gray_s.png");
                // メモ
                //不具合No1337（1094）対応
                c1FlexGrid1.Cols[14].Style.WordWrap = true;
                //c1FlexGrid1.Cols[25].Style.WordWrap = true;
                //C1.Win.C1FlexGrid.CellRange cr;
                //cr = c1FlexGrid1.GetCellRange(1, 7, c1FlexGrid1.Rows.Count - 1, 7);
                //cr.Image = Image.FromFile("folder_gray_s.png");

                //              C1.Win.C1FlexGrid.CellRange cr1;
                //            cr1 = c1FlexGrid1.GetCellRange(1, 2, c1FlexGrid1.Rows.Count - 1, 2);
                //          cr1.Image = Image.FromFile("file_presentation1.png");
            }

            //表示件数制御
            int viewnum = int.Parse(item_Hyoujikensuu.Text.Replace("件", ""));
            int startrow = (pagenum - 1) * viewnum;
            int addnum = ListData.Rows.Count - startrow;
            if (addnum > viewnum)
            {
                addnum = viewnum;
            }
            //データセット
            for (int r = 0; r < addnum; r++)
            {
                c1FlexGrid1.Rows.Add();
                for (int i = 0; i < c1FlexGrid1.Cols.Count - 1; i++)
                {

                    // 編集アイコンを表示する為に、1をセット
                    c1FlexGrid1[r + 1, 2] = "1";

                    // 進捗状況
                    //c1FlexGrid1[r + 1, 1] = ListData.Rows[startrow + r][31];
                    c1FlexGrid1[r + 1, 1] = ListData.Rows[startrow + r][32];

                    ////1列目のとき　進捗状況
                    //if (i == 0)
                    //{
                    //    String shinchoku = ListData.Rows[startrow + r][i].ToString();
                    //    String houkokuzumi = ListData.Rows[startrow + r][i + 1].ToString();
                    //    String shimekiri = ListData.Rows[startrow + r][i + 8].ToString();

                    //    //日付計算
                    //    DateTime shimekiribi;
                    //    DateTime today = DateTime.Today;

                    //    // 詳細画面の担当部所タブで締切日を空で更新可能の為、判定
                    //    if (shimekiri != "")
                    //    {
                    //        shimekiribi = DateTime.Parse(shimekiri);
                    //    }
                    //    else
                    //    {
                    //        // 値をセットしておかないと、後続で未割当となって参照できないので、ダミーで本日を入れる
                    //        shimekiribi = DateTime.Today;
                    //    }

                    //    //報告済み ソート8
                    //    if ("1".Equals(houkokuzumi))
                    //    {
                    //        c1FlexGrid1[r + 1, i + 1] = "8";
                    //    }
                    //    //報告済でない
                    //    else
                    //    {
                    //        //進捗が80：中止　のとき　ソート6
                    //        if ("80".Equals(shinchoku))
                    //        {
                    //            c1FlexGrid1[r + 1, i + 1] = "6";
                    //        }
                    //        // 進捗が70：二次検済 のとき ソート5
                    //        else if ("70".Equals(shinchoku))
                    //        {
                    //            c1FlexGrid1[r + 1, i + 1] = "5";
                    //        }
                    //        else if ("50".Equals(shinchoku) || "60".Equals(shinchoku))
                    //        {
                    //            // 担当者済み or 一次検済
                    //            c1FlexGrid1[r + 1, i + 1] = "7";
                    //        }
                    //        //締切日が過ぎている
                    //        else if (shimekiri != "")
                    //        {
                    //            if (shimekiribi < today)
                    //            {
                    //                c1FlexGrid1[r + 1, i + 1] = "1";
                    //            }
                    //            // 進捗が70：二次検済でなく、締切が今日から3日以内
                    //            else if (!"70".Equals(shinchoku) && shimekiribi <= today.AddDays(2))
                    //            {
                    //                c1FlexGrid1[r + 1, i + 1] = "2";
                    //            }
                    //            // 進捗が70：二次検済でなく、締切が今日から7日以内
                    //            else if (!"70".Equals(shinchoku) && shimekiribi <= today.AddDays(6))
                    //            {
                    //                c1FlexGrid1[r + 1, i + 1] = "3";
                    //            }
                    //        }
                    //        //それ以外
                    //        else
                    //        {
                    //            c1FlexGrid1[r + 1, i + 1] = "4";
                    //        }
                    //    }
                    //    //次のfor
                    //    continue;
                    //}//進捗状況end

                    //2列目　報告済　1列目のときに処理済み
                    if (i == 1)
                    {
                        continue;
                    }

                    //6列目　ファイルアイコン出すだけ（？）
                    if (i == 5)
                    {
                        //c1FlexGrid1[r + 1, i + 1] = "0";
                        //if (Directory.Exists(ListData.Rows[startrow + r][i].ToString()))
                        if (DirectoryExists(ListData.Rows[startrow + r][i].ToString()))
                        {
                            c1FlexGrid1[r + 1, i + 1] = "1";
                        }
                        else
                        {
                            c1FlexGrid1[r + 1, i + 1] = "0";
                        }
                        continue;
                    }

                    //c1FlexGrid1[行,列]値セット
                    c1FlexGrid1[r + 1, i + 1] = ListData.Rows[startrow + r][i];
                }//列for

            }
            c1FlexGrid1.AllowAddNew = false;
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
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
                string txt = e.Index > -1 ? ((ComboBox)sender).Items[e.Index].ToString() : ((ComboBox)sender).Text;
                e.Graphics.DrawString(txt, e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            }
            e.DrawFocusRectangle();
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
            get_data(int.Parse(Paging_now.Text));
            //set_data(int.Parse(Paging_now.Text));
            //set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
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

            //return ("" + ((DateTimePicker)cs[0]).Value.ToString() + "");
            DateTime Dt = ((DateTimePicker)cs[0]).Value.Date;
            return (Dt.ToString());
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

        // 検索期間の指定変更
        private void Kikansitei_TextChanged(object sender, EventArgs e)
        {
            //kensakukikanFlg = true;

            //if (item_FromTo.SelectedValue != null)
            //{
            //    // 締め日の選択を空に
            //    item_ShimekiriSentaku.SelectedValue = -1;
            //}

            //// 1:以前
            //if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() == "1")
            //{
            //    // 日付Fromが入力されている
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
            //            item_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //        }
            //        else
            //        {
            //            item_DateFrom.BackColor = Color.FromArgb(255, 255, 255);
            //        }
            //    }
            //    // Fromを消す
            //    item_DateFrom.CustomFormat = " ";

            //}
            //// 2:当日の場合
            //else if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() == "2")
            //{
            //    // 日付Fromが入力されている
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
            //            item_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //        }
            //        else
            //        {
            //            item_DateFrom.BackColor = Color.FromArgb(255, 255, 255);
            //            item_DateFrom.Value = item_DateTo.Value;
            //        }
            //    }
            //}
            //// 3:一週間
            //else if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() == "3")
            //{
            //    // 日付Fromが入力されている
            //    if (item_DateFrom.CustomFormat == "")
            //    {
            //        // Fromに6日を足した日をToにセット
            //        DateTime dateTime = item_DateFrom.Value;
            //        dateTime = dateTime.AddDays(6);
            //        // Toにコピー
            //        item_DateTo.Value = dateTime;
            //        item_DateTo.CustomFormat = "";
            //    }
            //    else
            //    {
            //        // Toが空だった場合
            //        if (item_DateTo.CustomFormat != "")
            //        {
            //            item_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //        }
            //        else
            //        {
            //            DateTime dateTime = item_DateTo.Value;
            //            // Toに6日を引いた日をFromにセット
            //            dateTime = dateTime.AddDays(-6);
            //            item_DateFrom.BackColor = Color.FromArgb(255, 255, 255);
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
            //DateTime dateTime;
            //// 未選択の場合
            //if (item_ShimekiriSentaku.Text == "")
            //{
            //    //検索期間・締切日の指定を空に
            //    item_FromTo.SelectedValue = -1;
            //    //item_FromTo.BackColor = Color.FromArgb(255, 255, 255);
            //    item_DateFrom.CustomFormat = " ";
            //    item_DateTo.CustomFormat = " ";
            //}

            //// 1:本日の締めは？
            //if (item_ShimekiriSentaku.SelectedValue != null && item_ShimekiriSentaku.SelectedValue.ToString() == "1" && kensakukikanFlg == false)
            //{
            //    // 本日日付を設定
            //    dateTime = DateTime.Today;
            //    item_DateFrom.Value = dateTime;
            //    item_DateFrom.CustomFormat = "";
            //    item_DateTo.Value = dateTime;
            //    item_DateTo.CustomFormat = "";
            //}
            //// 2:今週の締めは？
            //else if (item_ShimekiriSentaku.SelectedValue != null && item_ShimekiriSentaku.SelectedValue.ToString() == "2" && kensakukikanFlg == false)
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
            //    item_DateTo.CustomFormat = "";
            //}
            //// 3:来週の締めは？
            //else if (item_ShimekiriSentaku.SelectedValue != null && item_ShimekiriSentaku.SelectedValue.ToString() == "3" && kensakukikanFlg == false)
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
            //    item_DateTo.CustomFormat = "";
            //}

            //if (kensakukikanFlg)
            //{
            //    // 検索期間の指定を空に
            //    item_FromTo.SelectedValue = -1;
            //    //item_FromTo.BackColor = Color.FromArgb(255, 255, 255);
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

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            errorCheck_initialize();
            Boolean errorFlg = false;

            ////検索時必須チェック
            //if (search_required())
            //{
            //    get_data(1);
            //}
            errorFlg = !search_required();

            // 検索期間の指定に値が入っていた場合、入力チェック
            if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() != "")
            {
                // 値の再設定
                changeShimekiribi();

                errorFlg = errorCheck_Shimekiribi();
            }

            if (!errorFlg)
            {
                get_data(1);
            }

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        // 検索解除ボタン
        private void BtnClear_Click(object sender, EventArgs e)
        {
            errorCheck_initialize();
            ClearForm();
        }

        // 検索条件クリア
        private void ClearForm()
        {
            //検索条件初期化
            //今年度
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

            //年度オプション
            item_NendoOptionTounen.Checked = true;
            item_NendoOption3Nen.Checked = false;

            //調査区分
            //不具合No1337（1094）対応
            //item_ChousaKbnJibusho.Checked = false;
            //item_ChousaKbnShibuShibu.Checked = false;
            //item_ChousaKbnHonbuShibu.Checked = false;
            //item_ChousaKbnShibuHonbu.Checked = false;

            //発注者課名　特調番号
            item_HachushaKamei.Text = "";
            item_TokuchoNo.Text = "";

            // 日付From
            item_DateFrom.Text = "";
            item_DateFrom.CustomFormat = " ";
            label_DateFromTo.BackColor = Color.FromArgb(95, 158, 160);
            // 日付To
            item_DateTo.Text = "";
            item_DateTo.CustomFormat = " ";

            //部所　調査担当者
            item_ChousaTantouBusho.SelectedValue = UserInfos[2];
            item_MadoguchiBusho.SelectedIndex = -1;
            item_ChousaTantousha.Text = UserInfos[1];
            item_Taisho.SelectedIndex = 0;

            //業務名　管理番号　工事件名 メモ　部署備考
            //不具合No1337（1094）対応
            //item_Gyoumumei.Text = "";
            //不具合No1337（1094）対応
            //item_KanriBangou.Text = "";
            item_Koujikenmei.Text = "";
            item_Memo.Text = "";
            //不具合No1337（1094）対応
            //item_BushoBikou.Text = "";

            //本部単品
            //不具合No1337（1094）対応
            //item_HonbuTanpin.Checked = false;

            // 検索期間　締め日の選択　表示件数
            item_FromTo.SelectedIndex = -1;
            item_ShimekiriSentaku.SelectedIndex = -1;
            item_Hyoujikensuu.SelectedIndex = 1;

            //担当者状況　進捗状況
            item_TantouJoukyo.SelectedIndex = -1;
            item_Shintyokujyoukyo.SelectedIndex = -1;

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

        private void item_ChousaTantou_Icon_Click(object sender, EventArgs e)
        {
            //調査員プロンプト　窓口ミハルと同じ？TODO
            Popup_ChousainList form = new Popup_ChousainList();
            form.program = "madoguchi";

            if (!String.IsNullOrEmpty(item_ChousaTantouBusho.Text))
            {
                form.Busho = item_ChousaTantouBusho.SelectedValue.ToString();
            }

            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                //調査員名をセット
                item_ChousaTantousha.Text = form.ReturnValue[1];
            }
        }

        // Activeになったときの処理
        private void Jibun_Search_Activated(object sender, EventArgs e)
        {
            if (ReSearch)
            {
                get_data(1);
                ReSearch = false;
            }
        }

        // VIPS　20220301　課題管理表No1274(968)　DEL　「更新」処理追加　対応
        //private void button_Update_Click(object sender, EventArgs e)
        //{
        //    string methodName = "button_Update_Click";

        //    //更新ボタン押下処理
        //    set_error("", 0);

        //    var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
        //    using (var conn = new SqlConnection(connStr))
        //    {
        //        conn.Open();
        //        var cmd = conn.CreateCommand();
        //        SqlTransaction transaction = conn.BeginTransaction();
        //        cmd.Transaction = transaction;
        //        int updCount = 0;
        //        try
        //        {
        //            for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
        //            {
        //                //画面の担当者状況
        //                string gamenShinchoku = "";
        //                if (!String.IsNullOrEmpty(c1FlexGrid1.Rows[i][14].ToString()))
        //                {

        //                    gamenShinchoku = c1FlexGrid1.Rows[i][14].ToString();
        //                }

        //                //画面のメモ
        //                string gamenMemo = "";
        //                if (!String.IsNullOrEmpty(c1FlexGrid1.Rows[i][25].ToString()))
        //                {
        //                    gamenMemo = c1FlexGrid1.Rows[i][25].ToString();
        //                }


        //                //MadoguchiL1ChousaCDがない場合
        //                string chousaId = "";
        //                if (String.IsNullOrEmpty(c1FlexGrid1.Rows[i][30].ToString()))
        //                {
        //                    //メモも進捗も更新できないので次の行へ
        //                    continue;
        //                }
        //                else
        //                {
        //                    chousaId = c1FlexGrid1.Rows[i][30].ToString();
        //                }

        //                //差分フラグ
        //                Boolean sabun = false;

        //                //各行の14（担当者状況）と25（メモ）カラム目取得する
        //                DataTable dt = new DataTable();
        //                cmd.CommandText = "SELECT ISNULL(MadoguchiL1ChousaShinchoku,'') AS MadoguchiL1ChousaShinchoku, ISNULL(MadoguchiL1Memo,'') AS MadoguchiL1Memo " +
        //                    "FROM MadoguchiJouhouMadoguchiL1Chou " +
        //                    "WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][29].ToString() + " " +
        //                    "AND MadoguchiL1ChousaCD = " + chousaId + " ";

        //                var sda = new SqlDataAdapter(cmd);
        //                sda.Fill(dt);
        //                string shinchokuData = dt.Rows[0][0].ToString();
        //                string memoData = dt.Rows[0][1].ToString();
        //                //画面の担当者状況とデータの担当者状況の値が違う
        //                //if (!String.IsNullOrEmpty(shinchokuData) && !gamenShinchoku.Equals(shinchokuData)) 
        //                if (!gamenShinchoku.Equals(shinchokuData))
        //                {
        //                    sabun = true;
        //                }

        //                //画面のメモとデータのメモの値が違う
        //                //if (!String.IsNullOrEmpty(memoData) && !gamenMemo.Equals(memoData))
        //                if (!gamenMemo.Equals(memoData))
        //                {
        //                    sabun = true;
        //                }

        //                if (sabun)
        //                {
        //                    //各行の14（担当者状況）と25（メモ）カラム目 差分があるとき更新する
        //                    cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
        //                        "MadoguchiL1ChousaShinchoku = " + gamenShinchoku + " " +
        //                        ",MadoguchiL1Memo = N'" + gamenMemo + "' " +
        //                        ",MadoguchiL1AsteriaKoushinFlag = 1 " +
        //                        ",MadoguchiL1UpdateDate = SYSDATETIME() " +
        //                        ",MadoguchiL1UpdateUser = N'" + UserInfos[0] + "' " +
        //                        ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
        //                        "WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][29].ToString() + " " +
        //                        "AND MadoguchiL1ChousaCD = " + c1FlexGrid1.Rows[i][30].ToString() + " ";

        //                    cmd.ExecuteNonQuery();

        //                    //窓口情報の進捗
        //                    Boolean itijikenFlg = false;
        //                    string shinchoku = "10";
        //                    //中止以外の担当部所のデータを拾う
        //                    DataTable dt2 = new DataTable();
        //                    //cmd.CommandText = "SELECT MadoguchiL1ChousaShinchoku " +
        //                    cmd.CommandText = "SELECT min(MadoguchiL1ChousaShinchoku) " +
        //                        "FROM MadoguchiJouhouMadoguchiL1Chou " +
        //                        //"WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][29].ToString() + " " +
        //                        "WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][29].ToString() + " ";
        //                        //"AND MadoguchiL1ChousaShinchoku <> 80 ";

        //                    var sda2 = new SqlDataAdapter(cmd);
        //                    sda2.Fill(dt2);

        //                    for (int j = 0; j < dt2.Rows.Count; j++)
        //                    {
        //                        shinchoku = dt2.Rows[0][0].ToString();
        //                    }

        //                    //for (int j = 0; j < dt.Rows.Count; j++)
        //                    //{
        //                    //    //取得した進捗
        //                    //    shinchoku = int.Parse(dt.Rows[j][0].ToString());

        //                    //    //一次検査の場合
        //                    //    if (shinchoku == 60)
        //                    //    {
        //                    //        itijikenFlg = true;
        //                    //    }

        //                    //    //1行目
        //                    //    if (j == 0)
        //                    //    {
        //                    //        //一次検査の場合
        //                    //        if (shinchoku == 60)
        //                    //        {
        //                    //            //一時的に二次検済にする
        //                    //            shinchoku = 70;
        //                    //        }

        //                    //        //一次検済以外はこのまま
        //                    //    }
        //                    //    //2行目以降
        //                    //    else
        //                    //    {
        //                    //        //一次検査でないかつshinchokuが取得したshinchokuより大きい場合
        //                    //        if (shinchoku != 60 && shinchoku > int.Parse(dt.Rows[j][0].ToString()))
        //                    //        {
        //                    //            shinchoku = int.Parse(dt.Rows[j][0].ToString());
        //                    //        }
        //                    //    }//行数if

        //                    //}//進捗for end

        //                    ////一次検査があり、
        //                    //if (itijikenFlg)
        //                    //{
        //                    //    //shinchokuが二次検済より前の場合
        //                    //    if (shinchoku <= 50)
        //                    //    {
        //                    //        //特になし
        //                    //    }
        //                    //    else
        //                    //    {
        //                    //        //一次検査にする
        //                    //        shinchoku = 60;
        //                    //    }
        //                    //}

        //                    ////窓口情報の進捗状況を更新
        //                    //cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
        //                    //"MadoguchiShinchokuJoukyou = " + shinchoku + " " +
        //                    //",MadoguchiUpdateDate = SYSDATETIME()" +
        //                    //",MadoguchiUpdateUser = N'" + UserInfos[0] + "' " +
        //                    //",MadoguchiUpdateProgram = '" + pgmName + methodName + "' " +
        //                    //"WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][29].ToString() + " " +
        //                    //"AND MadoguchiHoukokuzumi <> 1 " +
        //                    ////"AND MadoguchiL1ChousaShinchoku <> 80 ";
        //                    //"AND MadoguchiShinchokuJoukyou <> 80 ";

        //                    // 窓口情報の進捗状況を更新・・・担当部所の最小の進捗で更新
        //                    cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
        //                    "MadoguchiShinchokuJoukyou = " + shinchoku + " " +
        //                    ",MadoguchiUpdateDate = SYSDATETIME()" +
        //                    ",MadoguchiUpdateUser = N'" + UserInfos[0] + "' " +
        //                    ",MadoguchiUpdateProgram = '" + pgmName + methodName + "' " +
        //                    "WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][29].ToString() + " ";
        //                    cmd.ExecuteNonQuery();


        //                    // 調査品目の進捗状況を更新・・・Gridで変更した進捗に更新
        //                    cmd.CommandText = "UPDATE ChousaHinmoku SET " +
        //                    "ChousaShinchokuJoukyou = " + gamenShinchoku + " " +
        //                    ",ChousaUpdateDate = SYSDATETIME()" +
        //                    ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +
        //                    ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
        //                    "WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][29].ToString() + " " +
        //                    "AND HinmokuChousainCD = " + c1FlexGrid1.Rows[i][12].ToString();

        //                    cmd.ExecuteNonQuery();

        //                    // ここでコミットしておかないと連携が動かない
        //                    transaction.Commit();

        //                    // 皇帝まもる連携
        //                    GlobalMethod.KouteiTantouBushoRenkei(c1FlexGrid1.Rows[i][29].ToString(), UserInfos[0], UserInfos[2]);

        //                    transaction = conn.BeginTransaction();
        //                    cmd.Transaction = transaction;

        //                    cmd.CommandText = "INSERT INTO T_HISTORY(" +
        //                    "H_DATE_KEY " +
        //                    ",H_NO_KEY " +
        //                    ",H_OPERATE_DT " +
        //                    ",H_OPERATE_USER_ID " +
        //                    ",H_OPERATE_USER_MEI " +
        //                    ",H_OPERATE_USER_BUSHO_CD " +
        //                    ",H_OPERATE_USER_BUSHO_MEI " +
        //                    ",H_OPERATE_NAIYO " +
        //                    ",H_ProgramName " +
        //                    ",H_TOKUCHOBANGOU " +
        //                    ",MadoguchiID " +
        //                    ",HistoryBeforeTantoubushoCD " +
        //                    ",HistoryBeforeTantoushaCD " +
        //                    ",HistoryAfterTantoubushoCD " +
        //                    ",HistoryAfterTantoushaCD " +
        //                    ")VALUES(" +
        //                    "SYSDATETIME() " +
        //                    ", " + GlobalMethod.getSaiban("HistoryID") + " " +
        //                    ",SYSDATETIME() " +
        //                    ",'" + UserInfos[0] + "' " +
        //                    ",N'" + UserInfos[1] + "' " +
        //                    ",'" + UserInfos[2] + "' " +
        //                    ",N'" + UserInfos[3] + "' " +
        //                    ",'自分大臣で更新を行いました。進捗状況:" + gamenShinchoku + "' " +
        //                    ",'" + pgmName + methodName + "' " +
        //                    ",'" + c1FlexGrid1.Rows[i][4].ToString() + "' " + // 特調番号
        //                    "," + c1FlexGrid1.Rows[i][29].ToString() + " " + // MadoguchiID
        //                    ",NULL " +
        //                    ",NULL " +
        //                    ",NULL " +
        //                    ",NULL " +
        //                    ")";
        //                cmd.ExecuteNonQuery();


        //                    updCount++;
        //                }//差分フラグtrue if end

        //            }//Grid for end

        //            //コミット
        //            transaction.Commit();

        //            //更新データ0件の場合
        //            if (updCount == 0)
        //            {
        //                //データの更新はありませんでした。
        //                set_error(GlobalMethod.GetMessage("I40002", ""));
        //            }
        //            else
        //            {
        //                //データを更新しました。
        //                set_error(GlobalMethod.GetMessage("I40001", ""));
        //                //履歴登録
        //                //cmd.CommandText = "INSERT INTO T_HISTORY(" +
        //                //    "H_DATE_KEY " +
        //                //    ",H_NO_KEY " +
        //                //    ",H_OPERATE_DT " +
        //                //    ",H_OPERATE_USER_ID " +
        //                //    ",H_OPERATE_USER_MEI " +
        //                //    ",H_OPERATE_USER_BUSHO_CD " +
        //                //    ",H_OPERATE_USER_BUSHO_MEI " +
        //                //    ",H_OPERATE_NAIYO " +
        //                //    ",H_ProgramName " +
        //                //    ",MadoguchiID " +
        //                //    ",HistoryBeforeTantoubushoCD " +
        //                //    ",HistoryBeforeTantoushaCD " +
        //                //    ",HistoryAfterTantoubushoCD " +
        //                //    ",HistoryAfterTantoushaCD " +
        //                //    ")VALUES(" +
        //                //    "SYSDATETIME() " +
        //                //    ", " + GlobalMethod.getSaiban("HistoryID") + " " +
        //                //    ",SYSDATETIME() " +
        //                //    ",'" + UserInfos[0] + "' " +
        //                //    ",N'" + UserInfos[1] + "' " +
        //                //    ",'" + UserInfos[2] + "' " +
        //                //    ",N'" + UserInfos[3] + "' " +
        //                //    ",'自分大臣で更新を行いました。' " +
        //                //    ",'" + pgmName + methodName + "' " +
        //                //    ",NULL " +
        //                //    ",NULL " +
        //                //    ",NULL " +
        //                //    ",NULL " +
        //                //    ",NULL " +
        //                //    ")";
        //            }


        //        }
        //        catch
        //        {
        //            transaction.Rollback();
        //            throw;
        //        }
        //        conn.Close();
        //    }
        //}

        //指定した締切日までの未提出分データを表示させる
        private void button_Miteishutu_Click(object sender, EventArgs e)
        {
            //メッセージクリア
            set_error("", 0);
            //締切日の色を戻す
            //label_DateFromTo.BackColor = Color.FromArgb(95, 158, 160);
            errorCheck_initialize();

            //未提出フラグ
            miteishutu = false;
            //締切日toが空のとき
            if (item_DateTo.CustomFormat == " ")
            {
                //締切日fromが空のとき
                if (item_DateFrom.CustomFormat == " ")
                {
                    //エラー
                    //label_DateFromTo.BackColor = Color.FromArgb(255, 204, 255);
                    label_DateTo.BackColor = errorBackColor;

                    //締切日を入力してください
                    set_error(GlobalMethod.GetMessage("E40001", ""));
                }
                //締切日fromが空でないとき
                else
                {
                    //未提出フラグ
                    miteishutu = true;

                    //fromの値をtoにコピーする
                    item_DateTo.CustomFormat = "";
                    item_DateTo.Value = item_DateFrom.Value;
                    item_DateFrom.Text = "";
                    item_DateFrom.CustomFormat = " ";

                    //検索条件の担当者状況を空にする
                    item_TantouJoukyo.SelectedIndex = -1;

                    //検索
                    if (search_required())
                    {
                        get_data(1);
                    }
                }
            }
            //締切日toが空でないとき
            else
            {
                //未提出フラグ
                miteishutu = true;

                //検索条件の担当者状況を空にする
                item_TantouJoukyo.SelectedIndex = -1;

                //検索をかける
                if (search_required())
                {
                    get_data(1);
                }
            }
            miteishutu = false;
        }

        // 日付系の値変更（CustomFormatを空文字にすると表示される）
        private void dateTimePicker_CloseUp(object sender, EventArgs e)
        {
            ((DateTimePicker)sender).CustomFormat = "";
        }

        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        private void c1FlexGrid1_KeyDownEdit(object sender, C1.Win.C1FlexGrid.KeyEditEventArgs e)
        {
            // フォルダアイコン
            if (e.Col == 25 && e.Row >= 1)
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

        //TOP　特調野郎
        private void button6_Click(object sender, EventArgs e)
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
        // 窓口ミハル
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

        // Gridの進捗状況切替
        private void c1FlexGrid1_AfterEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            DateTime dateTime = DateTime.Today;
            //不具合No1337（1094）対応 もともと13とハードコーディングされていたところを変数に変更した。
            int tanto_jokyo_index = 13;
            //その他の列インデックスも併せて変数に。いづれメンバのEnumに変更したい。
            int shinchoku_jokyo_index = 1;
            int shimekiribi_index = 8;
            int memo_index = 14;
            int madoguchi_id_index = 29;
            int madoguchi_chousa_id_index = 30;
            int hinmoku_chousain_id_index = 11;
            int tokucyo_no_index = 4;
            int houkokuzumi_index = 32;
            // 13:担当者状況変更 
            //// 14:担当者状況変更 
            //if (e.Row > 0 && (e.Col == 14))
            if (e.Row > 0 && (e.Col == tanto_jokyo_index))
            {
                // 1:報告済みの場合
                if ("1".Equals(c1FlexGrid1[e.Row, houkokuzumi_index].ToString()))
                {
                    // 報告済み
                    c1FlexGrid1.Rows[e.Row][shinchoku_jokyo_index] = "8";
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
                    if ("80".Equals(c1FlexGrid1[e.Row, tanto_jokyo_index].ToString()))
                    {
                        // 二次検証済み、または中止（中止）
                        c1FlexGrid1.Rows[e.Row][shinchoku_jokyo_index] = "6";
                    }
                    else if ("70".Equals(c1FlexGrid1[e.Row, tanto_jokyo_index].ToString()))
                    {
                        // 二次検証済み、または中止（二次検証済み）
                        c1FlexGrid1.Rows[e.Row][shinchoku_jokyo_index] = "5";
                    }
                    else if ("50".Equals(c1FlexGrid1[e.Row, tanto_jokyo_index].ToString()) || "60".Equals(c1FlexGrid1[e.Row, tanto_jokyo_index].ToString()))
                    {
                        // 担当者済み or 一次検済
                        c1FlexGrid1.Rows[e.Row][shinchoku_jokyo_index] = "7";
                    }
                    else if (c1FlexGrid1[e.Row, shimekiribi_index] != null)
                    {
                        try
                        {
                            dateTime = DateTime.Parse(c1FlexGrid1[e.Row, shimekiribi_index].ToString());
                            if (dateTime < DateTime.Today)
                            {
                                // 締切日経過
                                c1FlexGrid1.Rows[e.Row][shinchoku_jokyo_index] = "1";
                            }
                            else if (dateTime < DateTime.Today.AddDays(3))
                            {
                                // 締切日が3日以内、かつ2次検証が完了していない
                                c1FlexGrid1.Rows[e.Row][shinchoku_jokyo_index] = "2";
                            }
                            else if (dateTime < DateTime.Today.AddDays(7))
                            {
                                // 締切日が1週間以内、かつ2次検証が完了していない
                                c1FlexGrid1.Rows[e.Row][shinchoku_jokyo_index] = "3";
                            }
                            else
                            {
                                c1FlexGrid1.Rows[e.Row][shinchoku_jokyo_index] = "4";
                            }
                        }
                        catch
                        {
                            // 日付変換エラー
                            throw;
                        }
                    }

                }
            }// 14:担当者状況変更end

            // VIPS　20220301　課題管理表No1274(968)　ADD　「更新」処理追加　対応start
            string methodName = "c1FlexGrid1_AfterEdit";

            //更新ボタン押下処理
            set_error("", 0);

            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                SqlTransaction transaction = conn.BeginTransaction();
                cmd.Transaction = transaction;
                int updCount = 0;
                try
                {
                    int i = e.Row;

                    //画面の担当者状況
                    string gamenShinchoku = "";
                    if (!String.IsNullOrEmpty(c1FlexGrid1.Rows[i][tanto_jokyo_index].ToString()))
                    {
                        gamenShinchoku = c1FlexGrid1.Rows[i][tanto_jokyo_index].ToString();
                    }

                    //画面のメモ
                    string gamenMemo = "";
                    if (!String.IsNullOrEmpty(c1FlexGrid1.Rows[i][memo_index].ToString()))
                    {
                        gamenMemo = c1FlexGrid1.Rows[i][memo_index].ToString();
                    }

                    //MadoguchiL1ChousaCDがない場合
                    //メモも進捗も更新できないので、更新処理なし、「//更新データ0件の場合」の処理へ

                    string chousaId = "";

                    //MadoguchiL1ChousaCDがある場合、以下の処理
                    if (!String.IsNullOrEmpty(c1FlexGrid1.Rows[i][madoguchi_chousa_id_index].ToString()))
                    {
                        chousaId = c1FlexGrid1.Rows[i][madoguchi_chousa_id_index].ToString();
                        //差分フラグ
                        Boolean sabun = false;

                        //各行の14（担当者状況）と25（メモ）カラム目取得する
                        DataTable dt = new DataTable();
                        cmd.CommandText = "SELECT ISNULL(MadoguchiL1ChousaShinchoku,'') AS MadoguchiL1ChousaShinchoku, ISNULL(MadoguchiL1Memo,'') AS MadoguchiL1Memo " +
                            "FROM MadoguchiJouhouMadoguchiL1Chou " +
                            "WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][madoguchi_id_index].ToString() + " " +
                            "AND MadoguchiL1ChousaCD = " + chousaId + " ";

                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(dt);
                        string shinchokuData = dt.Rows[0][0].ToString();
                        string memoData = dt.Rows[0][1].ToString();
                        //画面の担当者状況とデータの担当者状況の値が違う
                        if (!gamenShinchoku.Equals(shinchokuData))
                        {
                            sabun = true;
                        }

                        //画面のメモとデータのメモの値が違う
                        if (!gamenMemo.Equals(memoData))
                        {
                            sabun = true;
                        }

                        if (sabun)
                        {
                            //各行の14（担当者状況）と25（メモ）カラム目 差分があるとき更新する
                            cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                                "MadoguchiL1ChousaShinchoku = " + gamenShinchoku + " " +
                                ",MadoguchiL1Memo = N'" + gamenMemo + "' " +
                                ",MadoguchiL1AsteriaKoushinFlag = 1 " +
                                ",MadoguchiL1UpdateDate = SYSDATETIME() " +
                                ",MadoguchiL1UpdateUser = N'" + UserInfos[0] + "' " +
                                ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
                                "WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][madoguchi_id_index].ToString() + " " +
                                "AND MadoguchiL1ChousaCD = " + c1FlexGrid1.Rows[i][madoguchi_chousa_id_index].ToString() + " ";

                            cmd.ExecuteNonQuery();

                            //窓口情報の進捗
                            string shinchoku = "10";
                            //中止以外の担当部所のデータを拾う
                            DataTable dt2 = new DataTable();
                            cmd.CommandText = "SELECT min(MadoguchiL1ChousaShinchoku) " +
                                "FROM MadoguchiJouhouMadoguchiL1Chou " +
                                "WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][madoguchi_id_index].ToString() + " ";

                            var sda2 = new SqlDataAdapter(cmd);
                            sda2.Fill(dt2);

                            for (int j = 0; j < dt2.Rows.Count; j++)
                            {
                                shinchoku = dt2.Rows[0][0].ToString();
                            }

                            // 窓口情報の進捗状況を更新・・・担当部所の最小の進捗で更新
                            cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                            "MadoguchiShinchokuJoukyou = " + shinchoku + " " +
                            ",MadoguchiUpdateDate = SYSDATETIME()" +
                            ",MadoguchiUpdateUser = N'" + UserInfos[0] + "' " +
                            ",MadoguchiUpdateProgram = '" + pgmName + methodName + "' " +
                            "WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][madoguchi_id_index].ToString() + " ";
                            cmd.ExecuteNonQuery();

                            // 調査品目の進捗状況を更新・・・Gridで変更した進捗に更新
                            cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                            "ChousaShinchokuJoukyou = " + gamenShinchoku + " " +
                            ",ChousaUpdateDate = SYSDATETIME()" +
                            ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +
                            ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                            "WHERE MadoguchiID = " + c1FlexGrid1.Rows[i][madoguchi_id_index].ToString() + " " +
                            "AND HinmokuChousainCD = " + c1FlexGrid1.Rows[i][hinmoku_chousain_id_index].ToString();

                            cmd.ExecuteNonQuery();

                            // ここでコミットしておかないと連携が動かない
                            transaction.Commit();

                            // 皇帝まもる連携
                            GlobalMethod.KouteiTantouBushoRenkei(c1FlexGrid1.Rows[i][madoguchi_id_index].ToString(), UserInfos[0], UserInfos[2]);

                            transaction = conn.BeginTransaction();
                            cmd.Transaction = transaction;

                            //No1419対応：メモ更新内容の出力
                            string memoBeforeChg = memoData;
                            string memoAfterChg = gamenMemo;
                            if (memoData.Length > 30)
                            {
                                memoBeforeChg = memoData.Substring(0, 30);
                            }
                            if (gamenMemo.Length > 30)
                            {
                                memoAfterChg = gamenMemo.Substring(0, 30);
                            }

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
                            ",HistoryBeforeTantoubushoCD " +
                            ",HistoryBeforeTantoushaCD " +
                            ",HistoryAfterTantoubushoCD " +
                            ",HistoryAfterTantoushaCD " +
                            ")VALUES(" +
                            "SYSDATETIME() " +
                            ", " + GlobalMethod.getSaiban("HistoryID") + " " +
                            ",SYSDATETIME() " +
                            ",'" + UserInfos[0] + "' " +
                            ",N'" + UserInfos[1] + "' " +
                            ",'" + UserInfos[2] + "' " +
                            ",N'" + UserInfos[3] + "' " +
                            //No1419対応：メモ更新内容の出力
                            //",'自分大臣で更新を行いました。進捗状況:" + gamenShinchoku + "' " +
                            ",'自分大臣で更新を行いました。進捗状況:" + gamenShinchoku + " 変更前メモ:" + memoBeforeChg + " 変更後メモ:" + memoAfterChg + "' " +
                            ",'" + pgmName + methodName + "' " +
                            ",N'" + c1FlexGrid1.Rows[i][tokucyo_no_index].ToString() + "' " + // 特調番号
                            "," + c1FlexGrid1.Rows[i][madoguchi_id_index].ToString() + " " + // MadoguchiID
                            ",NULL " +
                            ",NULL " +
                            ",NULL " +
                            ",NULL " +
                            ")";
                            cmd.ExecuteNonQuery();

                            updCount++;
                        }//差分フラグtrue if end


                        //コミット
                        transaction.Commit();
                    }

                    //更新データ0件の場合
                    if (updCount == 0)
                    {
                        //データの更新はありませんでした。
                        set_error(GlobalMethod.GetMessage("I40002", ""));
                    }
                    else
                    {
                        //データを更新しました。
                        set_error(GlobalMethod.GetMessage("I40001", ""));
                    }

                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
                conn.Close();
                // VIPS　20220301　課題管理表No1274(968)　ADD　「更新」処理追加　対応end
            }
        }

        private void button9_Click(object sender, EventArgs e)
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
                        // 7:受託調査業務工程表
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
                                report_data[0] = "0";
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
                                if (item_ChousaTantouBusho.Text != null && item_ChousaTantouBusho.Text != "")
                                {
                                    report_data[5] = item_ChousaTantouBusho.SelectedValue.ToString();
                                }
                                // 6.発注者名・課名
                                report_data[6] = item_HachushaKamei.Text;
                                // 7.特調番号
                                report_data[7] = item_TokuchoNo.Text;
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
                                if (item_TantouJoukyo.Text != null && item_TantouJoukyo.Text != "")
                                {
                                    // 工程表の出力が復活した場合、68,69はそのまま渡しても検索できないので注意
                                    report_data[10] = item_TantouJoukyo.SelectedValue.ToString();
                                }
                                // 11.窓口部所
                                if (item_MadoguchiBusho.Text != null && item_MadoguchiBusho.Text != "")
                                {
                                    report_data[11] = item_MadoguchiBusho.SelectedValue.ToString();
                                }
                                // 12.業務名称
                                //不具合No1337（1094）対応
                                //report_data[12] = item_Gyoumumei.Text;
                                report_data[12] = "";
                                // 13.管理番号
                                //不具合No1337（1094）対応
                                //report_data[13] = item_KanriBangou.Text;
                                report_data[13] = "";
                                // 14.検索期間の指定
                                report_data[14] = "0";
                                if (item_FromTo.Text != null && item_FromTo.Text != "")
                                {
                                    report_data[14] = item_FromTo.SelectedValue.ToString();
                                }
                                // 15.調査種別  自分大臣では未設定
                                report_data[15] = "0";
                                // 16.窓口担当者  自分大臣では未設定
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
                                //不具合No1337（1094）対応
                                //if (item_ChousaKbnJibusho.Checked)
                                //{
                                //    report_data[19] = "1";
                                //}
                                // 20.調査区分（支→支）
                                report_data[20] = "0";
                                //不具合No1337（1094）対応
                                //if (item_ChousaKbnShibuShibu.Checked)
                                //{
                                //    report_data[20] = "1";
                                //}
                                // 21.調査区分（本→支）
                                report_data[21] = "0";
                                //不具合No1337（1094）対応
                                //if (item_ChousaKbnHonbuShibu.Checked)
                                //{
                                //    report_data[21] = "1";
                                //}
                                // 22.調査区分（支→本）
                                report_data[22] = "0";
                                //不具合No1337（1094）対応
                                //if (item_ChousaKbnShibuHonbu.Checked)
                                //{
                                //    report_data[22] = "1";
                                //}
                                // 23.調査品目  自分大臣では未設定
                                // 24.進捗状況
                                report_data[24] = "0";
                                if (item_Shintyokujyoukyo.Text != null && item_Shintyokujyoukyo.Text != "")
                                {
                                    report_data[24] = item_Shintyokujyoukyo.SelectedValue.ToString();
                                }
                                // 25.本部単品
                                report_data[25] = "0";
                                //不具合No1337（1094）対応
                                //if (item_HonbuTanpin.Checked)
                                //{
                                //    report_data[25] = "1";
                                //}
                                // 26.調査担当者名
                                report_data[26] = item_ChousaTantousha.Text;
                                // 27.メモ
                                report_data[27] = item_Memo.Text;

                                string[] result = GlobalMethod.InsertMadoguchiReportWork(8, UserInfos[0], report_data, "KouteiHyo");

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
                        // 9:調査状況一覧
                        else if (Dt.Rows[0][0].ToString() == "7")
                        {
                            // string[]
                            // 23個分先に用意
                            string[] report_data = new string[23] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };

                            // 0.部所CD
                            report_data[0] = UserInfos[2];
                            // 1.登録年度
                            report_data[1] = item_Nendo.SelectedValue.ToString();
                            // 2.登録年度オプション
                            if (item_NendoOptionTounen.Checked)
                            {
                                report_data[2] = "1";   // 当年度
                            }
                            else
                            {
                                report_data[2] = "2";   // 3年以内
                            }
                            // 3.調査担当部所CD
                            if (item_ChousaTantouBusho.Text != null && item_ChousaTantouBusho.Text != "")
                            {
                                report_data[3] = item_ChousaTantouBusho.SelectedValue.ToString();
                            }
                            // 4.発注者名・課名
                            report_data[4] = item_HachushaKamei.Text;
                            // 5.特調番号
                            report_data[5] = item_TokuchoNo.Text;
                            // 6.締切日from
                            report_data[6] = "null";
                            if (item_DateFrom.CustomFormat == "")
                            {
                                report_data[6] = "'" + item_DateFrom.Text + "'";
                            }
                            // 7.締切日to
                            report_data[7] = "null";
                            if (item_DateTo.CustomFormat == "")
                            {
                                report_data[7] = "'" + item_DateTo.Text + "'";
                            }
                            // 8.担当者状況
                            report_data[8] = "0";
                            if (item_TantouJoukyo.Text != null && item_TantouJoukyo.Text != "")
                            {
                                // 工程表の出力が復活した場合、68,69はそのまま渡しても検索できないので注意
                                report_data[8] = item_TantouJoukyo.SelectedValue.ToString();
                            }
                            // 9.窓口部所
                            if (item_MadoguchiBusho.Text != null && item_MadoguchiBusho.Text != "")
                            {
                                report_data[9] = item_MadoguchiBusho.SelectedValue.ToString();
                            }
                            // 10.業務名称
                            //不具合No1337（1094）対応
                            //report_data[10] = item_Gyoumumei.Text;
                            report_data[10] = "";
                            // 11.管理番号
                            //不具合No1337（1094）対応
                            //report_data[11] = item_Gyoumumei.Text;    //←そもそも不具合のような。
                            report_data[11] = "";
                            // 12.工事件名
                            report_data[12] = item_Koujikenmei.Text;
                            // 13.調査区分（自部所）
                            report_data[13] = "0";
                            //不具合No1337（1094）対応
                            //if (item_ChousaKbnJibusho.Checked)
                            //{
                            //    report_data[13] = "1";
                            //}
                            // 14.調査区分（支→支）
                            report_data[14] = "0";
                            //不具合No1337（1094）対応
                            //if (item_ChousaKbnShibuShibu.Checked)
                            //{
                            //    report_data[14] = "1";
                            //}
                            // 15.調査区分（本→支）
                            report_data[15] = "0";
                            //不具合No1337（1094）対応
                            //if (item_ChousaKbnHonbuShibu.Checked)
                            //{
                            //    report_data[15] = "1";
                            //}
                            // 16.調査区分（支→本）
                            report_data[16] = "0";
                            //不具合No1337（1094）対応
                            //if (item_ChousaKbnShibuHonbu.Checked)
                            //{
                            //    report_data[16] = "1";
                            //}
                            // 17.進捗状況
                            report_data[17] = "0";
                            if (item_Shintyokujyoukyo.Text != null && item_Shintyokujyoukyo.Text != "")
                            {
                                report_data[17] = item_Shintyokujyoukyo.SelectedValue.ToString();
                            }
                            // 18.本部単品
                            report_data[18] = "0";
                            //不具合No1337（1094）対応
                            //if (item_HonbuTanpin.Checked)
                            //{
                            //    report_data[18] = "1";
                            //}
                            // 19.調査担当者名
                            report_data[19] = item_ChousaTantousha.Text;
                            // 20.メモ
                            report_data[20] = item_Memo.Text;
                            // 21.部所備考
                            //不具合No1337（1094）対応
                            //report_data[21] = item_BushoBikou.Text;
                            report_data[21] = "";
                            // 22.主副
                            report_data[22] = "0";
                            if (item_Taisho.Text != null && item_Taisho.Text != "")
                            {
                                report_data[22] = item_Taisho.SelectedValue.ToString();
                            }

                            string[] result = GlobalMethod.InsertMadoguchiReportWork(9, UserInfos[0], report_data, "ChousaJokyou");

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
                        //if (errorFLG == true)
                        //{
                        //    set_error("必須入力項目が入力されていません。");
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
            //label_DateFromTo.BackColor = Color.CadetBlue;
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
                //label_DateFromTo.BackColor = errorBackColor;
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
                    //label_DateFromTo.BackColor = errorBackColor;
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
            //    // height:595 → 1086・・・調査品目明細と合わせる
            //    // width:1864 → 3752
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid1.Height = 1086;
            //    c1FlexGrid1.Width = 3752;
            //}
            //else
            //{
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid1.Height = 595;
            //    c1FlexGrid1.Width = 1864;
            //}
            gridSizeChange();
        }
        private void gridSizeChange()
        {
            //if (btnGridSize.Text == "一覧拡大")
            //{
            //    // height:595 → 1086・・・調査品目明細と合わせる
            //    // width:1864 → 3752
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid1.Height = 1086;
            //    c1FlexGrid1.Width = 3752;
            //}
            //else
            //{
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid1.Height = 595;
            //    c1FlexGrid1.Width = 1864;
            //}
            string num = "";
            int bigHeight = 0;
            int bigWidth = 0;
            int smallHeight = 0;
            int smallWidth = 0;

            if (btnGridSize.Text == "一覧拡大")
            {
                num = GlobalMethod.GetCommonValue1("JIBUNDAIJIN_GRID_BIG_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out bigHeight);
                    if (bigHeight == 0)
                    {
                        bigHeight = 1086;
                    }
                }
                num = GlobalMethod.GetCommonValue1("JIBUNDAIJIN_GRID_BIG_WIDTH");
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
                num = GlobalMethod.GetCommonValue1("JIBUNDAIJIN_GRID_SMALL_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out smallHeight);
                    if (smallHeight == 0)
                    {
                        smallHeight = 595;
                    }
                }
                num = GlobalMethod.GetCommonValue1("JIBUNDAIJIN_GRID_SMALL_WIDTH");
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

        //ディレクトリ有無
        public static bool DirectoryExists(string path)
        {
            return TimeoutCore(() => Directory.Exists(path));
        }

        //タイムアウト処理部分 1秒タイムアウト
        private static bool TimeoutCore(Func<bool> existFunction)
        {
            int TimeoutSeconds = 1;
            var task = Task.Factory.StartNew(() => existFunction());
            return task.Wait(TimeoutSeconds * 1000) && task.Result;
        }

        private void button8_Click(object sender, EventArgs e)
        {

        }
    }
}

