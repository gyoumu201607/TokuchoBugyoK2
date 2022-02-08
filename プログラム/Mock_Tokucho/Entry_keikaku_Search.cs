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
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Runtime.InteropServices;
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
    public partial class Entry_keikaku_Search : Form
    {
        private string pgmName = "Entry_keikaku_Search";

        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();
        public string[] UserInfos;
        private string fileName;
        private int readCount = 0;
        private int errorId = 0;
        private string errorUser = "";

        // 処理待ち時間
        private static int waitTime = 3;

        // COMExceptionが発生したかのフラグ false:正常 true:エラー
        private static Boolean comexceptionFlg = false;
        private static int comexceptionCnt = 0;

        public Entry_keikaku_Search()
        {
            InitializeComponent();
        }

        private void Entry_keikaku_Search_Load(object sender, EventArgs e)
        {
            // ホイール制御
            this.src_1.MouseWheel += item_MouseWheel; // 売上年度
            this.src_4.MouseWheel += item_MouseWheel; // 計画部所支部
            this.src_8.MouseWheel += item_MouseWheel; // 契約区分
            this.src_12.MouseWheel += item_MouseWheel; // 案件数
            this.src_13.MouseWheel += item_MouseWheel; // 表示件数

            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            gridSizeChange();

            label3.Text = UserInfos[3] + "：" + UserInfos[1];
            label12.Text = GlobalMethod.GetCommonValue1("APL_VERSION");
            if (GlobalMethod.GetCommonValue1("BOOT_MODE") == "1")
            {
                label1.Text = GlobalMethod.GetCommonValue2("BOOT_MODE");
            }
            for (int i = 0; i < c1FlexGrid1.Cols.Count; i++)
            {
                c1FlexGrid1[1, i] = c1FlexGrid1[0, i];
            }

            // コントロールを初期化します。
            c1FlexGrid1.Styles.Normal.WordWrap = true;

            // 行ヘッダを作成します。
            c1FlexGrid1.Rows[0].AllowMerging = true;

            // 同一内容の４つのセルをマージします。
            C1.Win.C1FlexGrid.CellRange rng = c1FlexGrid1.GetCellRange(0, 14, 0, 16);
            rng.Data = "調査部";
            rng = c1FlexGrid1.GetCellRange(0, 17, 0, 19);
            rng.Data = "事業普及部";
            rng = c1FlexGrid1.GetCellRange(0, 20, 0, 22);
            rng.Data = "情報システム部";
            rng = c1FlexGrid1.GetCellRange(0, 23, 0, 25);
            rng.Data = "総合研究所";


            set_combo();
            set_defalt();
            get_date();
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));


            int PhaseFLG = GlobalMethod.GetIntroductionPhase();
            if (PhaseFLG <= 1)
            {
                button5.Visible = false;
            }
            if (UserInfos[4] != "2")
            {
                button3.Visible = false;
                button6.Visible = false;
                radioButton3.Visible = false;
                radioButton4.Visible = false;
            }
            //ソート項目にアイコンを設定
            C1.Win.C1FlexGrid.CellRange cr;
            Bitmap bmp1 = new Bitmap("Resource/Image/SortIconDefalt.png");
            Bitmap bmpSort = new Bitmap(bmp1, bmp1.Width / 6, bmp1.Height / 6);
            cr = c1FlexGrid1.GetCellRange(0, 3);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 4);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 5);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 6);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 7);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 8);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 9);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 10);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 11);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 12);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 13);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 27);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
        }

        //コンボボックスの内容を設定
        private void set_combo()
        {
            //売上年度
            String discript = "NendoSeireki";
            String value = "NendoID ";
            String table = "Mst_Nendo";

            // 
            //String where = "NendoID <= YEAR(GETDATE()) AND NendoID >= YEAR(GETDATE()) - 3 ORDER BY NendoID DESC";
            String where = "";
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            DataRow nendodr;
            if (combodt != null)
            {
                nendodr = combodt.NewRow();
                combodt.Rows.InsertAt(nendodr, 0);
            }
            src_1.DataSource = combodt;
            src_1.DisplayMember = "Discript";
            src_1.ValueMember = "Value";

            //計画部所支部（Grid）
            discript = "Mst_Busho.ShibuMei ";
            value = "DISTINCT Mst_Busho.BushoShibuCD ";
            table = "Mst_Busho";
            where = "GyoumuBushoCD < '999990' and BushoKeikakuHyoujiFlg = '1' " +
                "AND NOT GyoumuBushoCD LIKE '121%' " +
                //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' " +
                "AND BushoNewOld <= 1 AND ISNULL(BushoDeleteFlag,0) = 0 AND (BushoShibuCD IS NOT NULL OR KashoShibuCD IS NOT NULL) ";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            DataRow dr;
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }
            //グリッドのコンボボックス用リスト
            SortedList sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            c1FlexGrid1.Cols[5].DataMap = sl;

            //計画課所支部（Grid）
            discript = "BushokanriboKamei";
            value = "DISTINCT Mst_Busho.KashoShibuCD ";
            table = "Mst_Busho";
            where = "GyoumuBushoCD < '999990' and BushoKeikakuHyoujiFlg = '1' " +
                //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' " +
                "AND NOT GyoumuBushoCD LIKE '121%' " +
                "AND BushoNewOld <= 1 AND ISNULL(BushoDeleteFlag,0) = 0 AND (BushoShibuCD IS NOT NULL OR KashoShibuCD IS NOT NULL) ";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }
            //グリッドのコンボボックス用リスト
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            c1FlexGrid1.Cols[13].DataMap = sl;

            //契約区分
            discript = "GyoumuKubunHyouji";
            value = "GyoumuNarabijunCD";
            table = "Mst_GyoumuKubun";
            where = "GyoumuNarabijunCD < 100 ";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }
            src_8.DataSource = combodt;
            src_8.DisplayMember = "Discript";
            src_8.ValueMember = "Value";
            //グリッドのコンボボックス用リスト
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとSakuseiKubunをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            c1FlexGrid1.Cols[11].DataMap = sl;
        }

        private void set_combo_shibu(string nendo)
        {
            //計画支部(検索条件)
            string SelectedValue = "";
            if (src_4.Text != "")
            {
                SelectedValue = src_4.SelectedValue.ToString();
            }
            string discript = "Mst_Busho.ShibuMei + ' ' + ISNULL(KaMei, '') ";
            string value = "Mst_Busho.GyoumuBushoCD ";
            string table = "Mst_Busho";
            string where = "GyoumuBushoCD < '999990' and BushoKeikakuHyoujiFlg = '1' " +
                //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' " +
                "AND NOT GyoumuBushoCD LIKE '121%' " +
                "AND BushoNewOld <= 1 AND ISNULL(BushoDeleteFlag,0) = 0 AND (BushoShibuCD IS NOT NULL OR KashoShibuCD IS NOT NULL) ";
            //int FromNendo;
            //if (int.TryParse(nendo, out FromNendo))
            //{
            //    int ToNendo = int.Parse(nendo) + 1;
            //    if (src_3.Checked)
            //    {
            //        FromNendo -= 3;
            //    }
            //    where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
            //    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
            //}
            int FromNendo;
            if (int.TryParse(nendo, out FromNendo))
            {
                int ToNendo = int.Parse(nendo) + 1;
                if (src_3.Checked)
                {
                    //FromNendo -= 3;
                    ToNendo -= 2;
                }
                //where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
                where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' )";
            }
            where += " ORDER BY BushoEntoriNarabijun";
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            DataRow dr;
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }
            src_4.DataSource = combodt;
            src_4.DisplayMember = "Discript";
            src_4.ValueMember = "Value";
            if (SelectedValue != "")
            {
                src_4.SelectedValue = SelectedValue;
            }
        }

        private void set_defalt()
        {
            /*
            String discript = "NendoSeireki";
            String value = "NendoID ";
            String table = "Mst_Nendo";
            String where = "Nendo_Sdate <= GETDATE() AND Nendo_EDate >= GETDATE()";
            //コンボボックスデータ取得
            DataTable dt = GlobalMethod.getData(discript, value, table, where);
            if (dt != null)
            {
                src_1.SelectedValue = dt.Rows[0][0].ToString();
                //計画支部(検索条件)
                set_combo_shibu(dt.Rows[0][0].ToString());
            }
            else
            {
                src_1.SelectedValue = System.DateTime.Now.Year;
            }
            */
            src_1.SelectedValue = GlobalMethod.GetTodayNendo();
            //計画支部(検索条件)
            set_combo_shibu(GlobalMethod.GetTodayNendo());

            src_2.Checked = true;
            src_3.Checked = false;

            src_4.SelectedValue = UserInfos[2];
            src_5.Text = "";
            src_6.Text = "";
            src_7.Text = "";
            src_8.Text = "";
            src_9.Text = "";
            src_10.Text = "";
            src_11.Text = "";
            src_12.SelectedIndex = 0;
            src_13.SelectedIndex = 1;

            c1FlexGrid1.Rows.Count = 2;

            Paging_now.Text = "1";
            Paging_all.Text = "0";
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
        }

        // マウスポイントの下にある画像の取得ロジック
        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid1.HitTest(new Point(e.X, e.Y));

            if (hti.Column == 1 & hti.Row > 1)
            {
                var _row = hti.Row;
                var _col = hti.Column;
                Entry_keikaku_Detail form = new Entry_keikaku_Detail();
                form.KeikakuID = c1FlexGrid1.Rows[_row][2].ToString();
                form.UserInfos = this.UserInfos;
                //form.Show();
                form.Show(this);
                //this.Hide();
                //this.Show();
            }
        }

        // ヘッダーの案件ボタン
        private void button5_Click(object sender, EventArgs e)
        {
            Entry_Search form = new Entry_Search();
            form.UserInfos = this.UserInfos;
            form.Show();
            this.Close();
        }

        // 計画情報一覧取込
        private void button3_Click(object sender, EventArgs e)
        {
            Popup_Loading Loading = new Popup_Loading();
            Loading.StartPosition = FormStartPosition.CenterScreen;
            Loading.Show();

            errorId = 0;
            readCount = 0;
            errorUser = UserInfos[0] + "_" + DateTime.Today.ToString("yyyyMMdd");
            ErrorMessage.Text = "";
            ErrorMessage.Visible = false;

            //ファイル
            OpenFileDialog Dialog1 = new OpenFileDialog();
            Dialog1.InitialDirectory = @"C:";
            Dialog1.Title = "インポートするファイルを選択してください。";


            if (Dialog1.ShowDialog() == DialogResult.OK)
            {

                get_excel(Dialog1.FileName);
            }
            else
            {
                // E70039:ファイルが読み込まれていません。
                ErrorMessage.Text = GlobalMethod.GetMessage("E70039", "");
                ErrorMessage.Visible = true;
            }

            Dialog1.Dispose();
            Loading.Close();

            // バックグラウンドとなっているExcelプロセスをKILL
            excelProcessKill();
        }

        private void get_excel(string name)
        {
            string methodName = "get_excel";

            Application ExcelApp = null;
            Workbook wb = null;
            Worksheet ws = new Worksheet();

            try
            {
                //Excel取込
                ExcelApp = new Application();
                fileName = name;
                wb = getExcelFile(fileName, ExcelApp);

                if (wb == null)
                {
                    // E70039:ファイルが読み込まれていません。
                    ErrorMessage.Text = GlobalMethod.GetMessage("E70039", "");
                    ErrorMessage.Visible = true;
                    return;
                }

                //ファイル名から年度と部所支部名取得
                string[] del0 = { @"\" };
                string[] del1 = { @"_" };
                string[] del2 = { @"." };
                fileName = fileName.Split(del0, StringSplitOptions.RemoveEmptyEntries).Last();
                fileName = fileName.Replace(".xlsx", "");
                String[] faliNameSplit = fileName.Split(del1, StringSplitOptions.RemoveEmptyEntries);
                int result;

                String fileYear = "";
                String fileBusho = "";
                //if (faliNameSplit.Length >= 3)
                //{
                //    fileYear = faliNameSplit[1];
                //    fileBusho = faliNameSplit[2].Split(del2, StringSplitOptions.RemoveEmptyEntries)[0];
                //    int length = faliNameSplit[2].Length;
                //}
                for (int i = 0; faliNameSplit.Length > i; i++)
                {
                    switch (i)
                    {
                        case 0:     // ファイル名
                            break;
                        case 1:     // 年度
                            fileYear = faliNameSplit[i];
                            break;
                        case 2:     // 部所名
                            fileBusho = faliNameSplit[i];
                            break;
                        case 3:     // ？
                            break;
                        default:
                            break;
                    }
                }

                //シート取込 マスタ存在チェック
                Boolean check = checkMasta(wb, ws, out fileBusho, fileName, fileYear, fileBusho);

                //チェックがtrueの場合登録・更新処理
                if (check)
                {
                    //洗い替えの場合削除処理をする
                    if (radioButton4.Checked)
                    {
                        GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "計画情報の洗替を実施します。 ファイル名:" + fileName, pgmName + methodName, "");

                        deleteRecord(fileYear, fileBusho);
                    }
                    else
                    {
                        GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "計画情報の更新を実施します。 ファイル名:" + fileName, pgmName + methodName, "");
                    }
                    if (registrationKeikaku(wb, ws))
                    {
                        button6.Enabled = false;
                        button6.BackColor = Color.DarkGray;

                        // E70038:取込が完了しました。
                        ErrorMessage.Text = GlobalMethod.GetMessage("E70038", "");
                        ErrorMessage.Visible = true;

                        get_date();
                    }
                    else
                    {
                        // E70037:取込ファイルにエラーがありました。
                        ErrorMessage.Text = GlobalMethod.GetMessage("E70037", "");
                        ErrorMessage.Visible = true;

                        GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "計画情報の取込に失敗しました。 ファイル名:" + fileName, pgmName + methodName, "");

                    }

                }
                else
                {
                    // E70037:取込ファイルにエラーがありました。
                    ErrorMessage.Text = GlobalMethod.GetMessage("E70037", "");
                    ErrorMessage.Visible = true;
                    button6.Enabled = true;
                    button6.BackColor = Color.FromArgb(42, 78, 122);
                }

                //wb.Close(false);
                //ExcelApp.Quit();

            }
            finally
            {
                //Excelのオブジェクトを開放し忘れているとプロセスが落ちないため注意
                Marshal.ReleaseComObject(ws);
                wb.Close(false);
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
        }


        private Boolean checkMasta(Workbook wb, Worksheet ws, out string fileBusho, string FileName, string nendo, string busho)
        {
            Boolean checkFlag = true;
            dynamic xlSheet = null;
            fileBusho = "";
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


                    //Excelシートのインスタンスを作る
                    //インデックスは1始まりF
                    int u = wb.Sheets.Count;

                    //取込シートから部所支部名取得
                    //xlSheet = wb.Sheets[u - 1];
                    xlSheet = wb.Sheets["取り込みシート"];
                    //ws.Select(Type.Missing);
                    Excel.Range w_rgnName = xlSheet.Cells;
                    int getRowCount = xlSheet.UsedRange.Rows.Count;


                    DateTime tmp;
                    string ErrorMsg = "";
                    if (!DateTime.TryParse(w_rgnName[2, 3].Text, out tmp))
                    {
                        ErrorMsg += "・" + GlobalMethod.GetMessage("E70007", "");
                    }
                    //エクセルから部所名を取得
                    //fileBusho = w_rgnName[2, 4].Text;

                    if (ErrorMsg != "")
                    {
                        if (errorId == 0)
                        {
                            errorId = GlobalMethod.getSaiban("ErrorID");
                        }
                        if (readCount == 0)
                        {
                            //同じファイルで何度目のエラーか取得
                            readCount = GlobalMethod.getReadCount(errorUser);
                        }
                        GlobalMethod.InsertErrorTable(errorId, "KeikakuEXCEL", FileName, 2, ErrorMsg, errorUser, readCount, 1);
                        checkFlag = false;
                    }

                    // パフォーマンス改善
                    Worksheet worksheet = wb.Sheets["取り込みシート"];
                    worksheet.Select();

                    // Excel操作用の配列の作成
                    // 二次元配列の各次元の最小要素番号
                    int[] lower = { 1, 1 };
                    // 二次元配列の各次元の要素数 
                    int[] length = { getRowCount - 7, 40 };
                    object[,] InputObject = (object[,])Array.CreateInstance(typeof(object), length, lower);

                    // B8セルから、AO列の最後を取得する
                    Excel.Range InputRange = worksheet.Range[worksheet.Cells[8, 2], worksheet.Cells[getRowCount, 41]];
                    InputObject = (object[,])InputRange.Value;

                    String process = "";
                    String keikakuNo = "";

                    String nameBusho = "";
                    String nameKasho = "";
                    String keiyakuDivision = "";

                    int NendoFrom = 0;
                    int NendoTo = 0;

                    //int total = 0;
                    int haibun = 0;

                    long longChousabu = 0;      // 調査部配分　落札時見込額（税抜）
                    long longChousabuKei = 0;   // 調査部配分　見込額合計（税抜）

                    //Boolean totalFlg = false;
                    Boolean chousaCheckFlg = false;
                    Boolean PlanNoChk = true;

                    //for (int i = 8; getRowCount >= i; i++)
                    for (int i = 1; getRowCount - 7>= i; i++)
                    {
                        // 開始セル（B8＝[1, 1]）のため注意

                        ErrorMsg = "";
                        //計画番号が空欄の場合、以降空白行とみなし処理終了
                        //String process = w_rgnName[i, 2].Value;
                        //String keikakuNo = (string)w_rgnName[i, 5].Text;
                        //process = w_rgnName[i, 2].Value;
                        //keikakuNo = (string)w_rgnName[i, 5].Text;

                        process = "";
                        if (InputObject[i, 1] != null)
                        {
                            process = InputObject[i, 1].ToString();
                        }

                        //計画番号が空欄の場合、以降空白行とみなし処理終了
                        keikakuNo = "";
                        if (InputObject[i,4] != null)
                        {
                            keikakuNo = InputObject[i, 4].ToString();
                        }

                        if (String.IsNullOrEmpty(keikakuNo))
                        {
                            break;
                        }

                        //if (string.IsNullOrEmpty(keikakuNo))
                        //{
                        //    break;
                        //}

                        //必須チェック
                        //if (string.IsNullOrEmpty((string)w_rgnName[i, 2].Text)
                        //    || string.IsNullOrEmpty((string)w_rgnName[i, 4].Text)
                        //    || string.IsNullOrEmpty((string)w_rgnName[i, 5].Text)
                        //    || string.IsNullOrEmpty((string)w_rgnName[i, 11].Text)
                        //    || string.IsNullOrEmpty((string)w_rgnName[i, 12].Text)
                        //    || string.IsNullOrEmpty((string)w_rgnName[i, 13].Text))
                        if (InputObject[i, 1] == null
                        || InputObject[i, 3] == null
                        || InputObject[i, 4] == null
                        || InputObject[i, 10] == null
                        || InputObject[i, 11] == null
                        || InputObject[i, 12] == null)
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70003", "");
                        }


                        //処理区分ごとのチェック処理
                        ErrorMsg = processCheck(process, keikakuNo, ErrorMsg, nendo, busho);

                        //洗い替えで更新の場合エラー
                        //if (radioButton4.Checked && w_rgnName[i, 2].Value.Equals("更新"))
                        if (InputObject[i, 1] != null)
                        {
                            if (radioButton4.Checked && InputObject[i, 1].ToString().Equals("更新"))
                            {
                                ErrorMsg += "・" + GlobalMethod.GetMessage("E70044", "");
                            }

                        }

                        //部所支部名　課所支部名　契約区分
                        //String nameBusho = w_rgnName[i, 11].Text;
                        //String nameKasho = w_rgnName[i, 12].Text;
                        //String keiyakuDivision = w_rgnName[i, 13].Text;
                        //nameBusho = w_rgnName[i, 11].Text;
                        //nameKasho = w_rgnName[i, 12].Text;
                        //keiyakuDivision = w_rgnName[i, 13].Text;

                        //部所支部名
                        nameBusho = "";
                        if (InputObject[i, 10] != null)
                        {
                            nameBusho = InputObject[i, 10].ToString();
                        }

                        // 課所支部名
                        nameKasho = "";
                        if (InputObject[i, 11] != null)
                        {
                            nameKasho = InputObject[i, 11].ToString();
                        }

                        // 契約区分
                        keiyakuDivision = "";
                        if (InputObject[i, 12] != null)
                        {
                            keiyakuDivision = InputObject[i, 12].ToString();
                        }

                        //マスタ存在チェック
                        //int NendoFrom;
                        //if (!int.TryParse(w_rgnName[i, 4].Text, out NendoFrom))

                        NendoFrom = DateTime.Today.Year;
                        if (InputObject[i, 3] != null)
                        {
                            int.TryParse(InputObject[i, 3].ToString(), out NendoFrom);
                        }
                        //int NendoTo = NendoFrom + 1;
                        NendoTo = NendoFrom + 1;

                        if (!GlobalMethod.Check_Table(nameBusho, "ShibuMei", "Mst_Busho", "GyoumuBushoCD < '999990' and BushoKeikakuHyoujiFlg = '1' " +
                            "AND NOT GyoumuBushoCD LIKE '121%' " + 
                            //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' " +
                            "AND BushoNewOld <= 1 AND ISNULL(BushoDeleteFlag,0) = 0 AND (BushoShibuCD IS NOT NULL OR KashoShibuCD IS NOT NULL) " +
                            //"AND ( BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + NendoFrom + "/04/01' ) " +
                            //"AND ( BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + NendoTo + "/03/31' )"))
                            "AND ( BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + NendoTo + "/03/31' ) " +
                            "AND ( BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + NendoFrom + "/04/1' )"))
                        {
                            // E70004:部所支部名がマスタにありません。
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70004", "");
                        }
                        if (!GlobalMethod.Check_Table(nameKasho, "BushokanriboKamei", "Mst_Busho", "GyoumuBushoCD < '999990' and BushoKeikakuHyoujiFlg = '1' " +
                            "AND NOT GyoumuBushoCD LIKE '121%' " +
                            //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' " +
                            "AND BushoNewOld <= 1 AND ISNULL(BushoDeleteFlag,0) = 0 AND (BushoShibuCD IS NOT NULL OR KashoShibuCD IS NOT NULL) " +
                            //"AND ( BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + NendoFrom + "/04/01' ) " +
                            //"AND ( BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + NendoTo + "/03/31' )"))
                            "AND ( BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + NendoTo + "/03/31' ) " +
                            "AND ( BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + NendoFrom + "/04/1' )"))
                        {
                            // E70005:課所支部名がマスタにありません。
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70005", "");
                        }
                        if (!GlobalMethod.Check_Table(keiyakuDivision, "GyoumuKubun", "Mst_GyoumuKubun", ""))
                        {
                            // E70006:契約区分がマスタにありません。
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70006", "");
                        }

                        // ファイルの課所とのチェックはなし
                        //if (!nameBusho.Trim().Equals(fileBusho.Trim()))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70040", "");
                        //}


                        //if (!DateTime.TryParse(w_rgnName[i, 3].Text, out tmp))
                        //{
                        //    // E70008:計画登録日はYYYY/MM/DD形式で入力してください。
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70008", "");
                        //}
                        if (InputObject[i, 2] != null)
                        {
                            if (!DateTime.TryParse(InputObject[i, 2].ToString(), out tmp))
                            {
                                // E70008:計画登録日はYYYY/MM/DD形式で入力してください。
                                ErrorMsg += "・" + GlobalMethod.GetMessage("E70008", "");
                            }
                        }

                        //if (!Regex.IsMatch(w_rgnName[i, 4].Text, @"^[0-9]{4}$"))
                        //{
                        //    // E70009:売上年度はYYYY形式で入力してください。
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70009", "");
                        //}
                        if (InputObject[i, 3] != null)
                        {
                            if (!Regex.IsMatch(InputObject[i, 3].ToString(), @"^[0-9]{4}$"))
                            {
                                // E70009:売上年度はYYYY形式で入力してください。
                                ErrorMsg += "・" + GlobalMethod.GetMessage("E70009", "");
                            }

                        }

                        //調査部
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 16].Text) && !Regex.IsMatch(w_rgnName[i, 16].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70011", "");
                        //}
                        longChousabu = 0;
                        if (InputObject[i, 15] != null && !CheckMoney(InputObject[i, 15].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70011", "");
                        }
                        else
                        {
                            longChousabu = long.Parse(InputObject[i, 15].ToString());
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 17].Text) && !Regex.IsMatch(w_rgnName[i, 17].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70012", "");
                        //}
                        if (InputObject[i, 16] != null && !CheckMoney(InputObject[i, 16].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70012", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 18].Text) && !Regex.IsMatch(w_rgnName[i, 18].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70013", "");
                        //}
                        longChousabuKei = 0;
                        if (InputObject[i, 17] != null && !CheckMoney(InputObject[i, 17].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70013", "");
                        }
                        else
                        {
                            longChousabuKei = long.Parse(InputObject[i, 17].ToString());
                        }

                        //事業普及部
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 19].Text) && !Regex.IsMatch(w_rgnName[i, 19].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70014", "");
                        //}
                        if (InputObject[i, 18] != null && !CheckMoney(InputObject[i, 18].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70014", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 20].Text) && !Regex.IsMatch(w_rgnName[i, 20].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70015", "");
                        //}
                        if (InputObject[i, 19] != null && !CheckMoney(InputObject[i, 19].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70015", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 21].Text) && !Regex.IsMatch(w_rgnName[i, 21].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70016", "");
                        //}
                        if (InputObject[i, 20] != null && !CheckMoney(InputObject[i, 20].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70016", "");
                        }

                        //情報システム
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 22].Text) && !Regex.IsMatch(w_rgnName[i, 22].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70017", "");
                        //}
                        if (InputObject[i, 21] != null && !CheckMoney(InputObject[i, 21].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70017", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 23].Text) && !Regex.IsMatch(w_rgnName[i, 23].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70018", "");
                        //}
                        if (InputObject[i, 22] != null && !CheckMoney(InputObject[i, 22].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70018", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 24].Text) && !Regex.IsMatch(w_rgnName[i, 24].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70019", "");
                        //}
                        if (InputObject[i, 23] != null && !CheckMoney(InputObject[i, 23].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70019", "");
                        }

                        //総合研究所
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 25].Text) && !Regex.IsMatch(w_rgnName[i, 25].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70020", "");
                        //}
                        if (InputObject[i, 24] != null && !CheckMoney(InputObject[i, 24].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70020", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 26].Text) && !Regex.IsMatch(w_rgnName[i, 26].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70021", "");
                        //}
                        if (InputObject[i, 25] != null && !CheckMoney(InputObject[i, 25].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70021", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 27].Text) && !Regex.IsMatch(w_rgnName[i, 27].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70022", "");
                        //}
                        if (InputObject[i, 26] != null && !CheckMoney(InputObject[i, 26].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70022", "");
                        }

                        //見込み合計額
                        //if (!Regex.IsMatch(w_rgnName[i, 28].Text, @"^[\-]*[0-9\,\.]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70023", "");
                        //}
                        if (InputObject[i, 27] != null && !CheckMoney(InputObject[i, 27].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70023", "");
                        }

                        //業務別配分
                        //int total = 0;
                        //Boolean totalFlg = false;

                        // 調査部 業務別配分チェックフラグ true:チェックする false:チェックしない
                        // 29:資材調査 ～ 40:その他調査の合計が100かチェック
                        //Boolean chousaCheckFlg = false;
                        //chousaCheckFlg = false;

                        // 18:調査部配分 見込額合計（税抜）
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 18].Text) && Regex.IsMatch(w_rgnName[i, 18].Text, @"^[0-9]+$"))
                        //{
                        //    // 調査部配分 見込額合計（税抜） が 1以上の場合、調査部 業務別配分をチェックする
                        //    if (int.Parse(w_rgnName[i, 18].Text) > 0)
                        //    {
                        //        chousaCheckFlg = true;
                        //    }
                        //}
                        //if (InputObject[i, 17] != null && !CheckNumber(InputObject[i, 17].ToString()))
                        //{
                        //    // 調査部配分 見込額合計（税抜） が 1以上の場合、調査部 業務別配分をチェックする
                        //    if (int.Parse(InputObject[i, 17].ToString()) > 0)
                        //    {
                        //        chousaCheckFlg = true;
                        //    }
                        //}

                        //int haibun = 0;
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 29].Text) && !Regex.IsMatch(w_rgnName[i, 29].Text, @"^[0-9]+$"))
                        //{
                        //    // E70024:調査業務別　配分　資材調査は半角数字で入力してください。
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70024", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 29].Text, out haibun))
                        //    {
                        //        total += haibun;
                        //    }
                        //}
                        if (InputObject[i, 28] != null && !CheckNumber(InputObject[i, 28].ToString()))
                        {
                            // E70024:調査業務別　配分　資材調査は半角数字で入力してください。
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70024", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 30].Text) && !Regex.IsMatch(w_rgnName[i, 30].Text, @"^[0-9]+$"))
                        //{
                        //    // E70025:調査業務別　配分　営繕は半角数字で入力してください。
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70025", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 30].Text, out haibun))
                        //    {
                        //        total += haibun;
                        //    }
                        //}
                        if (InputObject[i, 29] != null && !CheckNumber(InputObject[i, 29].ToString()))
                        {
                            // E70025:調査業務別　配分　営繕は半角数字で入力してください。
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70025", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 31].Text) && !Regex.IsMatch(w_rgnName[i, 31].Text, @"^[0-9]+$"))
                        //{
                        //    // E70026:調査業務別　配分　機器類調査は半角数字で入力してください。
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70026", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 31].Text, out haibun))
                        //    {
                        //        total += haibun;
                        //        if (haibun > 0)
                        //        {
                        //            totalFlg = true;
                        //        }
                        //    }
                        //}
                        if (InputObject[i, 30] != null && !CheckNumber(InputObject[i, 30].ToString()))
                        {
                            // E70026:調査業務別　配分　機器類調査は半角数字で入力してください。
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70026", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 32].Text) && !Regex.IsMatch(w_rgnName[i, 32].Text, @"^[0-9]+$"))
                        //{
                        //    // E70027:調査業務別　配分　工事費調査は半角数字で入力してください。
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70027", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 32].Text, out haibun))
                        //    {
                        //        total += haibun;
                        //    }
                        //}
                        if (InputObject[i, 31] != null && !CheckNumber(InputObject[i, 31].ToString()))
                        {
                            // E70027:調査業務別　配分　工事費調査は半角数字で入力してください。
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70027", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 33].Text) && !Regex.IsMatch(w_rgnName[i, 33].Text, @"^[0-9]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70028", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 33].Text, out haibun))
                        //    {
                        //        total += haibun;
                        //    }
                        //}
                        if (InputObject[i, 32] != null && !CheckNumber(InputObject[i, 32].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70028", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 34].Text) && !Regex.IsMatch(w_rgnName[i, 34].Text, @"^[0-9]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70029", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 34].Text, out haibun))
                        //    {
                        //        total += haibun;
                        //    }
                        //}
                        if (InputObject[i, 33] != null && !CheckNumber(InputObject[i, 33].ToString()))
                        {
                        ErrorMsg += "・" + GlobalMethod.GetMessage("E70029", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 35].Text) && !Regex.IsMatch(w_rgnName[i, 35].Text, @"^[0-9]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70030", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 35].Text, out haibun))
                        //    {
                        //        total += haibun;
                        //    }
                        //}
                        if (InputObject[i, 34] != null && !CheckNumber(InputObject[i, 34].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70030", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 36].Text) && !Regex.IsMatch(w_rgnName[i, 36].Text, @"^[0-9]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70031", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 36].Text, out haibun))
                        //    {
                        //        total += haibun;
                        //    }
                        //}
                        if (InputObject[i, 35] != null && !CheckNumber(InputObject[i, 35].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70031", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 37].Text) && !Regex.IsMatch(w_rgnName[i, 37].Text, @"^[0-9]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70032", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 37].Text, out haibun))
                        //    {
                        //        total += haibun;
                        //    }
                        //}
                        if (InputObject[i, 36] != null && !CheckNumber(InputObject[i, 36].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70032", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 38].Text) && !Regex.IsMatch(w_rgnName[i, 38].Text, @"^[0-9]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70033", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 38].Text, out haibun))
                        //    {
                        //        total += haibun;
                        //    }
                        //}
                        if (InputObject[i, 37] != null && !CheckNumber(InputObject[i, 37].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70033", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 39].Text) && !Regex.IsMatch(w_rgnName[i, 39].Text, @"^[0-9]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70034", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 39].Text, out haibun))
                        //    //if (int.TryParse(InputObject[i, 39].ToString(), out haibun))
                        //    {
                        //        total += haibun;
                        //    }
                        //}
                        if (InputObject[i, 38] != null && !CheckNumber(InputObject[i, 38].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70034", "");
                        }
                        //if (!string.IsNullOrEmpty(w_rgnName[i, 40].Text) && !Regex.IsMatch(w_rgnName[i, 40].Text, @"^[0-9]+$"))
                        //{
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70035", "");
                        //}
                        //else
                        //{
                        //    if (int.TryParse(w_rgnName[i, 40].Text, out haibun))
                        //    {
                        //        total += haibun;
                        //    }
                        //}
                        if (InputObject[i, 39] != null && !CheckNumber(InputObject[i, 39].ToString()))
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70035", "");
                        }

                        int intGyomuHaibunKei = 0;
                        //業務配分合計
                        //if (!Regex.IsMatch(w_rgnName[i, 41].Text, @"^[0-9]+$"))
                        //{
                        //    // E70036:調査業務別　配分　合計は半角数字で入力してください。
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70036", "");
                        //}
                        if (InputObject[i, 40] != null && !CheckNumber(InputObject[i, 40].ToString()))
                        {
                            // E70036:調査業務別　配分　合計は半角数字で入力してください。
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70036", "");
                        }
                        else
                        {
                            intGyomuHaibunKei = int.Parse(InputObject[i, 40].ToString());
                        }


                        //if (totalFlg == true && total != 100)
                        //{
                        //    // E70045:調査業務別　配分の合計が100になるように入力してください。
                        //    ErrorMsg += "・" + GlobalMethod.GetMessage("E70045", "");
                        //}
                        // 
                        if (chousaCheckFlg == true)
                        {
                            //if (int.TryParse(w_rgnName[i, 41].Text, out haibun))
                            //{
                            //    if (haibun != 100)
                            //    {
                            //        // E70045:調査業務別　配分の合計が100になるように入力してください。
                            //        ErrorMsg += "・" + GlobalMethod.GetMessage("E70045", "");
                            //    }
                            //}
                            if (InputObject[i, 40] != null)
                            {
                                if (int.TryParse(InputObject[i, 40].ToString(), out haibun))
                                {
                                    if (haibun != 100)
                                    {
                                        // E70045:調査業務別　配分の合計が100になるように入力してください。
                                        ErrorMsg += "・" + GlobalMethod.GetMessage("E70045", "");
                                    }
                                }
                            }
                        }
                        if (longChousabuKei != 0 || intGyomuHaibunKei != 0)
                        {
                            if (longChousabuKei > 0 && intGyomuHaibunKei != 100)
                            {
                                // E70045:調査業務別　配分の合計が100になるように入力してください。
                                ErrorMsg += "・" + GlobalMethod.GetMessage("E70045", "");
                            }

                            if (longChousabuKei == 0 && intGyomuHaibunKei > 0)
                            {
                                // 落札見込額と変更見込額で相殺して0円になったケースを除くため、落札見込額が0かチェックする
                                if (longChousabu == 0)
                                {
                                    // E70045:調査業務別　配分を入力するには、業務別配分の入力をしてください。
                                    ErrorMsg += "・" + GlobalMethod.GetMessage("E70056", "");
                                }
                            }

                        }

                        // 一次保留
                        ////調査部見込み額が0より上でない場合、
                        //if (string.IsNullOrEmpty(w_rgnName[i, 18].Text) || "0".Equals(w_rgnName[i, 18].Text))
                        //{
                        //    //調査業務別配分が入ってるとエラー
                        //    if ((!string.IsNullOrEmpty(w_rgnName[i, 29].Text) && !"0".Equals(w_rgnName[i, 29].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 30].Text) && !"0".Equals(w_rgnName[i, 30].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 31].Text) && !"0".Equals(w_rgnName[i, 31].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 32].Text) && !"0".Equals(w_rgnName[i, 32].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 33].Text) && !"0".Equals(w_rgnName[i, 33].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 34].Text) && !"0".Equals(w_rgnName[i, 34].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 35].Text) && !"0".Equals(w_rgnName[i, 35].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 36].Text) && !"0".Equals(w_rgnName[i, 36].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 37].Text) && !"0".Equals(w_rgnName[i, 37].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 38].Text) && !"0".Equals(w_rgnName[i, 38].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 39].Text) && !"0".Equals(w_rgnName[i, 39].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 40].Text) && !"0".Equals(w_rgnName[i, 40].Text))
                        //        || (!string.IsNullOrEmpty(w_rgnName[i, 41].Text) && !"0".Equals(w_rgnName[i, 41].Text))
                        //        )
                        //    {
                        //        ErrorMsg += "・調査業務別配分を入力するには、調査部配分を割り振ってください";
                        //    }
                        //}
                        //計画番号形式チェック
                        //Boolean PlanNoChk = true;
                        PlanNoChk = true;
                        if (keikakuNo.Length != 10)
                        {
                            PlanNoChk = false;
                        }
                        else
                        {
                            if (!GlobalMethod.Check_Table(keikakuNo.Substring(0, 2), "BushoShibuCD", "Mst_Busho", ""))
                            {
                                PlanNoChk = false;
                            }
                            if (!GlobalMethod.Check_Table(keikakuNo.Substring(2, 1), "JigyoubuHeadCD", "Mst_Jigyoubu", ""))
                            {
                                PlanNoChk = false;
                            }
                            //if (!Regex.IsMatch(keikakuNo.Substring(3, 2), @"^[0-9]+$"))
                            if (!CheckNumber(keikakuNo.Substring(3, 2)))
                            {
                                PlanNoChk = false;
                            }
                            if (keikakuNo.Substring(5, 1) != "-")
                            {
                                PlanNoChk = false;
                            }
                            if (keikakuNo.Substring(6, 1) != "P")
                            {
                                PlanNoChk = false;
                            }
                            //if (!Regex.IsMatch(keikakuNo.Substring(7, 3), @"^[0-9]+$"))
                            if (!CheckNumber(keikakuNo.Substring(7, 3)))
                            {
                                PlanNoChk = false;
                            }
                        }
                        if (!PlanNoChk)
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70041", "");
                        }

                        if (ErrorMsg != "")
                        {
                            if (errorId == 0)
                            {
                                errorId = GlobalMethod.getSaiban("ErrorID");
                            }
                            if (readCount == 0)
                            {
                                //同じファイルで何度目のエラーか取得
                                readCount = GlobalMethod.getReadCount(errorUser);
                            }
                            // エクセルの行番号と合わせるため、カウントに7を加算する。
                            GlobalMethod.InsertErrorTable(errorId, "KeikakuEXCEL", FileName, i + 7, ErrorMsg, errorUser, readCount, 1);
                            checkFlag = false;
                        }

                    }
                    transaction.Commit();
                    conn.Close();
                }
            }
            catch (Exception)
            {
                checkFlag = false;
            }
            finally
            {
                xlSheet = null;
            }
            return checkFlag;
        }

        private Boolean deleteRecord(String fileYear, String fileBusho)
        {
            Boolean endFlag = true;
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                //洗い替えのため、削除処理
                conn.Open();
                var cmd = conn.CreateCommand();
                SqlTransaction transaction = conn.BeginTransaction();
                cmd.Transaction = transaction;

                try
                {
                    //cmd.CommandText = "SELECT KeikakuBangou FROM KeikakuJouhou WHERE KeikakuUriageNendo = '" + fileYear + "' AND KeikakuBushoShibuMei = '" + fileBusho + "' ";
                    // 年度で洗い替え
                    cmd.CommandText = "SELECT KeikakuBangou FROM KeikakuJouhou WHERE KeikakuUriageNendo = '" + fileYear + "'  ";
                    DataTable dt = new DataTable();

                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);
                    string KeikakuBangou = "";
                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            if (i != 0)
                            {
                                KeikakuBangou += ", ";
                            }
                            KeikakuBangou += "N'" + dt.Rows[i][0].ToString() + "'";
                        }
                        cmd.CommandText = "UPDATE AnkenJouhou SET AnkenKeikakuBangou = '' , AnkenKeikakuAnkenMei = '' WHERE AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC IN (" + KeikakuBangou + ")";
                        cmd.ExecuteNonQuery();
                    }


                    cmd.CommandText = "DELETE FROM KeikakuJouhou " +
                        "WHERE KeikakuUriageNendo = '" + fileYear + "' ";
                    //"AND KeikakuBushoShibuMei = '" + fileBusho + "' ";

                    cmd.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    endFlag = false;
                }
                finally
                {
                    conn.Close();
                }
            }
            return endFlag;
        }


        private string processCheck(String processing, String keikakuNo, string ErrorMsg, string nendo, string busho)
        {
            //処理区分別のチェック　DBから情報を取得
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                var comboDt = new DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                "KeikakuID ,KeikakuUriageNendo ,KeikakuBushoShibuMei " +
                "FROM KeikakuJouhou " +
                "WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(keikakuNo, 0, 0) + "' ";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(comboDt);

                //登録の場合　登録済みの計画番号を指定した場合のエラー
                if ("登録".Equals(processing) && comboDt.Rows.Count == 1)
                {
                    if (radioButton4.Checked)
                    {
                        //if (comboDt.Rows[0][1].ToString().Equals(nendo) && comboDt.Rows[0][2].ToString().Equals(busho))
                        if (comboDt.Rows[0][1].ToString().Equals(nendo))
                        {

                        }
                        else
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70042", "");
                        }
                    }
                    else
                    {
                        ErrorMsg += "・" + GlobalMethod.GetMessage("E70042", "");
                    }
                }
                else if ("更新".Equals(processing) && comboDt.Rows.Count == 0)
                {
                    //更新の場合　対象データがなかったらエラー
                    ErrorMsg += "・" + GlobalMethod.GetMessage("E70043", "");
                }
                else if ("削除".Equals(processing))
                {
                    // 洗替で削除はエラー
                    if (radioButton4.Checked)
                    {
                        ErrorMsg += "・" + GlobalMethod.GetMessage("E70044", "");
                    }
                    else
                    {
                        // 削除対象が存在しない場合、エラー
                        if (comboDt.Rows.Count == 0)
                        {
                            ErrorMsg += "・" + GlobalMethod.GetMessage("E70043", "");
                        }
                    }
                }

                return ErrorMsg;
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
                if (ExcelApp.Ready == true || waitTime >= num)
                {
                    break;
                }
                Thread.Sleep(1000);
                num += 1;
            }
            ExcelApp.Visible = false;
            return wb;
        }

        private Boolean registrationKeikaku(Workbook wb, Worksheet ws)
        {
            string methodName = ".registrationKeikaku";
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

                    try
                    {
                        //Excelシートのインスタンスを作る
                        //インデックスは1始まり
                        int u = wb.Sheets.Count;

                        //取込シートから部所支部名取得
                        dynamic xlSheet = null;
                        //xlSheet = wb.Sheets[u - 1];
                        //ws = wb.Sheets[u - 1];
                        xlSheet = wb.Sheets["取り込みシート"];
                        ws = wb.Sheets["取り込みシート"];
                        //ws.Select(Type.Missing);
                        //w_rgnName = xlSheet.Cells;
                        w_rgnName = xlSheet.UsedRange;

                        int getRowCount = ws.UsedRange.Rows.Count;

                        string InsertSQL;

                        string Sakuseibi = (string)w_rgnName[2, 3].Text;
                        string KashoShibuMei = (string)w_rgnName[3, 3].Text;

                        // パフォーマンス改善
                        Worksheet worksheet = wb.Sheets["取り込みシート"];
                        worksheet.Select();

                        // Excel操作用の配列の作成
                        // 二次元配列の各次元の最小要素番号
                        int[] lower = { 1, 1 };
                        // 二次元配列の各次元の要素数 
                        int[] length = { getRowCount - 7, 40 };
                        object[,] InputObject = (object[,])Array.CreateInstance(typeof(object), length, lower);

                        // B8セルから、AO列の最後を取得する
                        Excel.Range InputRange = worksheet.Range[worksheet.Cells[8, 2], worksheet.Cells[getRowCount, 41]];
                        InputObject = (object[,])InputRange.Value;

                        String procces = "";
                        String planNo = "";
                        String nameBusho = "";

                        int saibanPlanNo = 0;

                        // 登録用SQL
                        String InsertSQLText = GetInsertKeikakuJouhou();

                        ////登録更新処理
                        ////for (int i = 8; getRowCount >= i; i++)
                        //for (int i = 1; getRowCount -7>= i; i++)
                        //{

                        //    //処理区分 部所支部名
                        //    //String procces = w_rgnName[i, 2].Text;
                        //    //String planNo = (string)w_rgnName[i, 5].Text;
                        //    //String nameBusho = w_rgnName[i, 11].Text;
                        //    //procces = w_rgnName[i, 2].Text;
                        //    //planNo = (string)w_rgnName[i, 5].Text;
                        //    //nameBusho = w_rgnName[i, 11].Text;
                        //    procces = "";
                        //    if (InputObject[i, 1] != null)
                        //    {
                        //        procces = InputObject[i, 1].ToString();
                        //    }
                        //    planNo = "";
                        //    if (InputObject[i, 4] != null)
                        //    {
                        //        planNo = InputObject[i, 4].ToString();
                        //    }
                        //    nameBusho = "";
                        //    if (InputObject[i, 10] != null)
                        //    {
                        //        nameBusho = InputObject[i, 10].ToString();
                        //    }

                        //    //planNoが取得できない場合、以降空白行とみなし処理終了
                        //    if (String.IsNullOrEmpty(planNo))
                        //    {
                        //        break;
                        //    }

                        //    GlobalMethod.outputLogger(pgmName + methodName, "計画IDの取得", planNo.ToString(), UserInfos[1]);

                        //    // 計画IDの取得
                        //    var KeikakuDt = new DataTable();
                        //    cmd.CommandText = "SELECT KeikakuID FROM KeikakuJouhou WHERE KeikakuBangou = '" + planNo + "'";
                        //    var sda = new SqlDataAdapter(cmd);
                        //    sda.Fill(KeikakuDt);
                        //    saibanPlanNo = 0;
                        //    if (KeikakuDt != null && KeikakuDt.Rows.Count > 0)
                        //    {
                        //        int.TryParse(KeikakuDt.Rows[0][0].ToString(), out saibanPlanNo);
                        //    }

                        //    //計画情報登録更新
                        //    if (procces.Equals("登録"))
                        //    {
                        //        GlobalMethod.outputLogger(pgmName + methodName, "計画情報の登録", planNo.ToString(), UserInfos[1]);

                        //        //処理区分が登録のときは新規登録
                        //        //int saibanPlanNo = GlobalMethod.getSaiban("KeikakuID");
                        //        saibanPlanNo = GlobalMethod.getSaiban("KeikakuID");

                        //        //計画情報登録
                        //        //cmd.CommandText = "INSERT INTO KeikakuJouhou(" +
                        //        //    "KeikakuID " +
                        //        //    ",KeikakuFileMei " +                //ファイル名
                        //        //    ",KeikakuSakuseibi " +              //作成日
                        //        //    ",KeikakuSakuseiKashoShibuMei " +   //課所支部名
                        //        //    ",KeikakuTourokubi " +              //計画登録日
                        //        //    ",KeikakuUriageNendo " +            //売上年度
                        //        //    ",KeikakuBangou " +                 //計画番号
                        //        //    ",KeikakuZenkaiKeikakuBangou " +
                        //        //    ",KeikakuZenkaiAnkenBangou " +
                        //        //    ",KeikakuZenkaiJutakuBangou " +
                        //        //    ",KeikakuZenkaiHachuushaMeiKaMei " +
                        //        //    ",KeikakuZenkaiGyoumuMei " +
                        //        //    ",KeikakuBushoShibuCD " +
                        //        //    ",KeikakuBushoShibuMei " +
                        //        //    ",KeikakuKashoShibuCD " +
                        //        //    ",KeikakuKashoShibuMei " +
                        //        //    ",KeikakuGyoumuKubun " +
                        //        //    ",KeikakuGyoumuKubunMei " +
                        //        //    ",KeikakuHachuushaMeiKaMei " +
                        //        //    ",KeikakuAnkenMei " +
                        //        //    ",KeikakuRakusatsumikomigaku " +
                        //        //    ",KeikakuHenkoumikomigaku " +
                        //        //    ",KeikakuMikomigaku " +
                        //        //    ",KeikakuRakusatsumikomigakuJF " +
                        //        //    ",KeikakuHenkoumikomigakuJF " +
                        //        //    ",KeikakuMikomigakuJF " +
                        //        //    ",KeikakuRakusatsumikomigakuJ " +
                        //        //    ",KeikakuHenkoumikomigakuJ " +
                        //        //    ",KeikakuMikomigakuJ " +
                        //        //    ",KeikakuRakusatsumikomigakuK " +
                        //        //    ",KeikakuHenkoumikomigakuK " +
                        //        //    ",KeikakuMikomigakuK " +
                        //        //    ",KeikakuMikomigakuGoukei " +
                        //        //    ",KeikakuShizaiChousa " +
                        //        //    ",KeikakuEizen " +
                        //        //    ",KeikakuKikiruiChousa " +
                        //        //    ",KeikakuKoujiChousahi " +
                        //        //    ",KeikakuSanpaiChousa " +
                        //        //    ",KeikakuHokakeChousa " +
                        //        //    ",KeikakuShokeihiChousa " +
                        //        //    ",KeikakuGenkaBunseki " +
                        //        //    ",KeikakuKijunsakusei " +
                        //        //    ",KeikakuKoukyouRoumuhi " +
                        //        //    ",KeikakuRoumuhiKoukyouigai " +
                        //        //    ",KeikakuSonotaChousabu " +
                        //        //    ",KeikakuHaibunGoukei " +
                        //        //    ",KeikakuAnkensu " +
                        //        //    ",KeikakuCreateDate " +
                        //        //    ",KeikakuCreateUser " +
                        //        //    ",KeikakuCreateProgram " +
                        //        //    ",KeikakuUpdateDate " +
                        //        //    ",KeikakuUpdateUser " +
                        //        //    ",KeikakuUpdateProgram " +
                        //        //    ",KeikakuDeleteFlag " +
                        //        //    ")VALUES(" +
                        //        //    saibanPlanNo +
                        //        //", '" + GlobalMethod.ChangeSqlText(fileName, 0, 100) + "' " + //ファイル名
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[2, 3].Text, 0, 0) + "' " + //作成日
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[3, 3].Text, 0, 50) + "' " + //課所支部名
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 3].Text, 0, 0) + "' " + //計画登録日
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 4].Text, 0, 4) + "' " + //売上年度
                        //        //", '" + GlobalMethod.ChangeSqlText(planNo, 0, 10) + "' " + //計画番号
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 6].Text, 0, 10) + "' " + //前年度計画番号
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 7].Text, 0, 40) + "' " + //前年度案件番号
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 8].Text, 0, 40) + "' " + //前年度受託番号
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 9].Text, 0, 50) + "' " + //前年度発注者名・課名
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 10].Text, 0, 100) + "' " + //前年度業務名称
                        //        //", '" + DiscriptToValue(w_rgnName[i, 11].Text, "ShibuMei", "BushoShibuCD", "Mst_Busho") + "' " + //部所支部CD
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 11].Text, 0, 50) + "' " + //部所支部名
                        //        //", '" + DiscriptToValue(w_rgnName[i, 12].Text, "BushokanriboKamei", "KashoShibuCD", "Mst_Busho") + "' " + //課所支部CD
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 12].Text, 0, 50) + "' " + //課所支部名
                        //        //", '" + DiscriptToValue(w_rgnName[i, 13].Text, "GyoumuKubun", "GyoumuNarabijunCD", "Mst_GyoumuKubun") + "' " + //契約区分CD
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 13].Text, 0, 50) + "' " + //契約区分
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 14].Text, 0, 100) + "' " + //発注者名・課名
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 15].Text, 0, 150) + "' " + //計画案件名
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 16].Text, 0, 0) + "' " + //調査部配分　落札時見込額（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 17].Text, 0, 0) + "' " + //調査部配分　変更見込額（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 18].Text, 0, 0) + "' " + //調査部配分　見込額合計（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 19].Text, 0, 0) + "' " + //事業普及部配分　落札時見込額（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 20].Text, 0, 0) + "' " + //事業普及部配分　変更見込額（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 21].Text, 0, 0) + "' " + //事業普及部配分　見込額合計（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 22].Text, 0, 0) + "' " + //情報システム部配分　落札時見込額（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 23].Text, 0, 0) + "' " + //情報システム部配分　変更見込額（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 24].Text, 0, 0) + "' " + //情報システム部配分　見込額合計（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 25].Text, 0, 0) + "' " + //総合研究所　落札時見込額（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 26].Text, 0, 0) + "' " + //総合研究所　変更見込額（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 27].Text, 0, 0) + "' " + //総合研究所　見込額合計（税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 28].Text, 0, 0) + "' " + //見込額総合計(税抜）
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 29].Text, 0, 0) + "' " + //調査業務別　配分　資材調査
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 30].Text, 0, 0) + "' " + //調査業務別　配分　営繕
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 31].Text, 0, 0) + "' " + //調査業務別　配分　機器類調査
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 32].Text, 0, 0) + "' " + //調査業務別　配分　工事費調査
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 33].Text, 0, 0) + "' " + //調査業務別　配分　産廃調査
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 34].Text, 0, 0) + "' " + //調査業務別　配分　歩掛調査
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 35].Text, 0, 0) + "' " + //調査業務別　配分　諸経費調査
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 36].Text, 0, 0) + "' " + //調査業務別　配分　原価分析調査
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 37].Text, 0, 0) + "' " + //調査業務別　配分　基準作成改訂
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 38].Text, 0, 0) + "' " + //調査業務別　配分　公共労費調査
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 39].Text, 0, 0) + "' " + //調査業務別　配分　労務費公共以外
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 40].Text, 0, 0) + "' " + //調査業務別　配分　その他調査部
                        //        //", '" + GlobalMethod.ChangeSqlText(w_rgnName[i, 41].Text, 0, 0) + "' " + //調査業務別　配分　合計
                        //        //", " + GetAnkensu(planNo) + " " + //案件数
                        //        //", SYSDATETIME()  " + //作成日時
                        //        //", '" + UserInfos[0] + "' " + //作成ユーザ
                        //        //", '計画情報取込' " + //作成機能
                        //        //", SYSDATETIME()  " + //更新日時
                        //        //", '" + UserInfos[0] + "' " + //更新ユーザ
                        //        //", '計画情報取込' " + //更新機能
                        //        //", '0' " + //削除フラグ
                        //        //")";

                        //        cmd.CommandText = InsertSQLText
                        //                        + saibanPlanNo
                        //                        + ", '" + GlobalMethod.ChangeSqlText(fileName, 0, 100) + "' "//ファイル名
                        //                        + ", '" + GlobalMethod.ChangeSqlText(Sakuseibi, 0, 0) + "' " //作成日
                        //                        + ", '" + GlobalMethod.ChangeSqlText(KashoShibuMei, 0, 50) + "' " //課所支部名
                        //                        + ", '" + ChangeSQLString(InputObject[i, 2], 0, 0) + "' " //計画登録日
                        //                        + ", '" + ChangeSQLString(InputObject[i, 3], 0, 4) + "' " //売上年度
                        //                        + ", '" + ChangeSQLString(planNo, 0, 10) + "' " //計画番号
                        //                        + ", '" + ChangeSQLString(InputObject[i, 5], 0, 10) + "' " //前年度計画番号
                        //                        + ", '" + ChangeSQLString(InputObject[i, 6], 0, 40) + "' " //前年度案件番号
                        //                        + ", '" + ChangeSQLString(InputObject[i, 7], 0, 40) + "' " //前年度受託番号
                        //                        + ", '" + ChangeSQLString(InputObject[i, 8], 0, 50) + "' " //前年度発注者名・課名
                        //                        + ", '" + ChangeSQLString(InputObject[i, 9], 0, 100) + "' " //前年度業務名称
                        //                        + ", '" + DiscriptToValue(ChangeSQLString(InputObject[i, 10], 0, 50), "ShibuMei", "BushoShibuCD", "Mst_Busho") + "' " //部所支部CD
                        //                        + ", '" + ChangeSQLString(InputObject[i, 10], 0, 50) + "' " //部所支部名
                        //                        + ", '" + DiscriptToValue(ChangeSQLString(InputObject[i, 11], 0, 50), "BushokanriboKamei", "KashoShibuCD", "Mst_Busho") + "' " //課所支部CD
                        //                        + ", '" + ChangeSQLString(InputObject[i, 11], 0, 50) + "' " //課所支部名
                        //                        + ", '" + DiscriptToValue(ChangeSQLString(InputObject[i, 12], 0, 50), "GyoumuKubun", "GyoumuNarabijunCD", "Mst_GyoumuKubun") + "' " //契約区分CD
                        //                        + ", '" + ChangeSQLString(InputObject[i, 12], 0, 50) + "' " //契約区分
                        //                        + ", '" + ChangeSQLString(InputObject[i, 13], 0, 100) + "' " //発注者名・課名
                        //                        + ", '" + ChangeSQLString(InputObject[i, 14], 0, 150) + "' " //計画案件名
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 15], 0, 0) + "' " //調査部配分　落札時見込額（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 16], 0, 0) + "' " //調査部配分　変更見込額（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 17], 0, 0) + "' " //調査部配分　見込額合計（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 18], 0, 0) + "' " //事業普及部配分　落札時見込額（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 19], 0, 0) + "' " //事業普及部配分　変更見込額（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 20], 0, 0) + "' " //事業普及部配分　見込額合計（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 21], 0, 0) + "' " //情報システム部配分　落札時見込額（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 22], 0, 0) + "' " //情報システム部配分　変更見込額（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 23], 0, 0) + "' " //情報システム部配分　見込額合計（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 24], 0, 0) + "' " //総合研究所　落札時見込額（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 25], 0, 0) + "' " //総合研究所　変更見込額（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 26], 0, 0) + "' " //総合研究所　見込額合計（税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 27], 0, 0) + "' " //見込額総合計(税抜）
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 28], 0, 0) + "' " //調査業務別　配分　資材調査
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 29], 0, 0) + "' " //調査業務別　配分　営繕
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 30], 0, 0) + "' " //調査業務別　配分　機器類調査
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 31], 0, 0) + "' " //調査業務別　配分　工事費調査
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 32], 0, 0) + "' " //調査業務別　配分　産廃調査
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 33], 0, 0) + "' " //調査業務別　配分　歩掛調査
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 34], 0, 0) + "' " //調査業務別　配分　諸経費調査
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 35], 0, 0) + "' " //調査業務別　配分　原価分析調査
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 36], 0, 0) + "' " //調査業務別　配分　基準作成改訂
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 37], 0, 0) + "' " //調査業務別　配分　公共労費調査
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 38], 0, 0) + "' " //調査業務別　配分　労務費公共以外
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 39], 0, 0) + "' " //調査業務別　配分　その他調査部
                        //                        + ", '" + ChangeSQLInt(InputObject[i, 40], 0, 0) + "' " //調査業務別　配分　合計
                        //                        + ", " + GetAnkensu(planNo) + " " //案件数
                        //                        + ", SYSDATETIME()  " //作成日時
                        //                        + ", '" + UserInfos[0] + "' " //作成ユーザ
                        //                        + ", '計画情報取込' " //作成機能
                        //                        + ", SYSDATETIME()  " //更新日時
                        //                        + ", '" + UserInfos[0] + "' " //更新ユーザ
                        //                        + ", '計画情報取込' " //更新機能
                        //                        + ", '0' " //削除フラグ
                        //                        + ")";

                        //        cmd.ExecuteNonQuery();

                        //        GlobalMethod.outputLogger(pgmName + methodName, "変更履歴の登録", planNo.ToString(), UserInfos[1]);
                        //        // 共通メソッドを使うと、コミットされてしまうので、個別に記載する
                        //        //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "計画情報を登録しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName, "");
                        //        cmd.CommandText = Insert_History_SQLText("計画情報を登録しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName);
                        //        cmd.ExecuteNonQuery();

                        //    }
                        //    else if (procces.Equals("更新"))
                        //    {
                        //        GlobalMethod.outputLogger(pgmName + methodName, "計画情報の更新", planNo.ToString(), UserInfos[1]);

                        //        //処理区分が更新のときは更新処理 
                        //        cmd.CommandText = "UPDATE KeikakuJouhou SET " +
                        //            "KeikakuSakuseibi                   = '" + GlobalMethod.ChangeSqlText(Sakuseibi, 0, 0) + "' " +
                        //            ",KeikakuFileMei                    = '" + GlobalMethod.ChangeSqlText(fileName, 0, 100) + "' " +
                        //            ",KeikakuSakuseiKashoShibuMei       = '" + GlobalMethod.ChangeSqlText(KashoShibuMei, 0, 50) + "' " +
                        //            ",KeikakuTourokubi                  = '" + ChangeSQLString(InputObject[i, 2], 0, 0) + "' " +
                        //            ",KeikakuUriageNendo                = '" + ChangeSQLString(InputObject[i, 3], 0, 4) + "' " +
                        //            ",KeikakuZenkaiKeikakuBangou        = '" + ChangeSQLString(InputObject[i, 5], 0, 10) + "' " +
                        //            ",KeikakuZenkaiAnkenBangou          = '" + ChangeSQLString(InputObject[i, 6], 0, 40) + "' " +
                        //            ",KeikakuZenkaiJutakuBangou         = '" + ChangeSQLString(InputObject[i, 7], 0, 40) + "' " +
                        //            ",KeikakuZenkaiHachuushaMeiKaMei    = '" + ChangeSQLString(InputObject[i, 8], 0, 100) + "' " +
                        //            ",KeikakuZenkaiGyoumuMei            = '" + ChangeSQLString(InputObject[i, 9], 0, 150) + "' " +
                        //            ",KeikakuBushoShibuCD               = '" + DiscriptToValue(ChangeSQLString(InputObject[i, 10],0,50), "ShibuMei", "BushoShibuCD", "Mst_Busho") + "' " +
                        //            ",KeikakuBushoShibuMei              = '" + ChangeSQLString(InputObject[i, 10], 0, 50) + "' " +
                        //            ",KeikakuKashoShibuCD               = '" + DiscriptToValue(ChangeSQLString(InputObject[i, 11], 0, 50), "BushokanriboKamei", "KashoShibuCD", "Mst_Busho") + "' " +
                        //            ",KeikakuKashoShibuMei              = '" + ChangeSQLString(InputObject[i, 11], 0, 50) + "' " +
                        //            ",KeikakuGyoumuKubun                = '" + DiscriptToValue(ChangeSQLString(InputObject[i, 12], 0, 50), "GyoumuKubun", "GyoumuNarabijunCD", "Mst_GyoumuKubun") + "' " +
                        //            ",KeikakuGyoumuKubunMei             = '" + ChangeSQLString(InputObject[i, 12], 0, 50) + "' " +
                        //            ",KeikakuHachuushaMeiKaMei          = '" + ChangeSQLString(InputObject[i, 13], 0, 100) + "' " +
                        //            ",KeikakuAnkenMei                   = '" + ChangeSQLString(InputObject[i, 14], 0, 150) + "' " +
                        //            ",KeikakuRakusatsumikomigaku        = '" + ChangeSQLInt(InputObject[i, 15], 0, 0) + "' " +
                        //            ",KeikakuHenkoumikomigaku           = '" + ChangeSQLInt(InputObject[i, 16], 0, 0) + "' " +
                        //            ",KeikakuMikomigaku                 = '" + ChangeSQLInt(InputObject[i, 17], 0, 0) + "' " +
                        //            ",KeikakuRakusatsumikomigakuJF      = '" + ChangeSQLInt(InputObject[i, 18], 0, 0) + "' " +
                        //            ",KeikakuHenkoumikomigakuJF         = '" + ChangeSQLInt(InputObject[i, 19], 0, 0) + "' " +
                        //            ",KeikakuMikomigakuJF               = '" + ChangeSQLInt(InputObject[i, 20], 0, 0) + "' " +
                        //            ",KeikakuRakusatsumikomigakuJ       = '" + ChangeSQLInt(InputObject[i, 21], 0, 0) + "' " +
                        //            ",KeikakuHenkoumikomigakuJ          = '" + ChangeSQLInt(InputObject[i, 22], 0, 0) + "' " +
                        //            ",KeikakuMikomigakuJ                = '" + ChangeSQLInt(InputObject[i, 23], 0, 0) + "' " +
                        //            ",KeikakuRakusatsumikomigakuK       = '" + ChangeSQLInt(InputObject[i, 24], 0, 0) + "' " +
                        //            ",KeikakuHenkoumikomigakuK          = '" + ChangeSQLInt(InputObject[i, 25], 0, 0) + "' " +
                        //            ",KeikakuMikomigakuK                = '" + ChangeSQLInt(InputObject[i, 26], 0, 0) + "' " +
                        //            ",KeikakuMikomigakuGoukei           = '" + ChangeSQLInt(InputObject[i, 27], 0, 0) + "' " +
                        //            ",KeikakuShizaiChousa               = '" + ChangeSQLInt(InputObject[i, 28], 0, 0) + "' " +
                        //            ",KeikakuEizen                      = '" + ChangeSQLInt(InputObject[i, 29], 0, 0) + "' " +
                        //            ",KeikakuKikiruiChousa              = '" + ChangeSQLInt(InputObject[i, 30], 0, 0) + "' " +
                        //            ",KeikakuKoujiChousahi              = '" + ChangeSQLInt(InputObject[i, 31], 0, 0) + "' " +
                        //            ",KeikakuSanpaiChousa               = '" + ChangeSQLInt(InputObject[i, 32], 0, 0) + "' " +
                        //            ",KeikakuHokakeChousa               = '" + ChangeSQLInt(InputObject[i, 33], 0, 0) + "' " +
                        //            ",KeikakuShokeihiChousa             = '" + ChangeSQLInt(InputObject[i, 34], 0, 0) + "' " +
                        //            ",KeikakuGenkaBunseki               = '" + ChangeSQLInt(InputObject[i, 35], 0, 0) + "' " +
                        //            ",KeikakuKijunsakusei               = '" + ChangeSQLInt(InputObject[i, 36], 0, 0) + "' " +
                        //            ",KeikakuKoukyouRoumuhi             = '" + ChangeSQLInt(InputObject[i, 37], 0, 0) + "' " +
                        //            ",KeikakuRoumuhiKoukyouigai         = '" + ChangeSQLInt(InputObject[i, 38], 0, 0) + "' " +
                        //            ",KeikakuSonotaChousabu             = '" + ChangeSQLInt(InputObject[i, 39], 0, 0) + "' " +
                        //            ",KeikakuHaibunGoukei               = '" + ChangeSQLInt(InputObject[i, 40], 0, 0) + "' " +
                        //            ",KeikakuAnkensu                    = " + GetAnkensu(planNo) + " " +
                        //            ",KeikakuUpdateDate                 = SYSDATETIME() " +
                        //            ",KeikakuUpdateUser                 = '" + UserInfos[0] + "' " +
                        //            ",KeikakuUpdateProgram              = '計画情報取込' " +
                        //            ",KeikakuDeleteFlag                 = '0' " +
                        //            "WHERE KeikakuBangou                = '" + planNo + "' ";

                        //        cmd.ExecuteNonQuery();

                        //        GlobalMethod.outputLogger(pgmName + methodName, "変更履歴の登録", planNo.ToString(), UserInfos[1]);
                        //        // 共通メソッドを使うと、コミットされてしまうので、個別に記載する
                        //        //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "計画情報を更新しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName, "");
                        //        cmd.CommandText = Insert_History_SQLText("計画情報を更新しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName);
                        //        cmd.ExecuteNonQuery();
                        //    }
                        //    else if (procces.Equals("削除"))
                        //    {

                        //        GlobalMethod.outputLogger(pgmName + methodName, "案件情報の更新", planNo.ToString(), UserInfos[1]);

                        //        // 案件情報から計画情報のクリア
                        //        cmd.CommandText = "UPDATE AnkenJouhou SET"
                        //                        + "  AnkenKeikakuBangou = ''"                       // 計画番号
                        //                        + ", AnkenKeikakuAnkenMei = ''"                     // 計画案件名
                        //                        + " WHERE AnkenKeikakuBangou = '" + planNo + "'"
                        //                        ;
                        //        cmd.ExecuteNonQuery();

                        //        GlobalMethod.outputLogger(pgmName + methodName, "計画情報の削除", planNo.ToString(), UserInfos[1]);

                        //        // 計画情報の削除
                        //        cmd.CommandText = "DELETE FROM KeikakuJouhou"
                        //                        + " WHERE KeikakuBangou = '" + planNo + "'"
                        //                        ;
                        //        cmd.ExecuteNonQuery();

                        //        GlobalMethod.outputLogger(pgmName + methodName, "変更履歴の登録", planNo.ToString(), UserInfos[1]);
                        //        // 共通メソッドを使うと、コミットされてしまうので、個別に記載する
                        //        //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "計画情報を削除しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName, "");
                        //        cmd.CommandText = Insert_History_SQLText("計画情報を削除しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName);
                        //        cmd.ExecuteNonQuery();

                        //    }
                        //}

                        // 後でちゃんと直す。登録、更新、削除を分けて処理する
                        // 登録処理
                        for (int i = 1; getRowCount - 7 >= i; i++)
                        {

                            //処理区分 部所支部名
                            procces = "";
                            if (InputObject[i, 1] != null)
                            {
                                procces = InputObject[i, 1].ToString();
                            }
                            planNo = "";
                            if (InputObject[i, 4] != null)
                            {
                                planNo = InputObject[i, 4].ToString();
                            }
                            nameBusho = "";
                            if (InputObject[i, 10] != null)
                            {
                                nameBusho = InputObject[i, 10].ToString();
                            }

                            //planNoが取得できない場合、以降空白行とみなし処理終了
                            if (String.IsNullOrEmpty(planNo))
                            {
                                break;
                            }

                            //// 計画IDの取得
                            //var KeikakuDt = new DataTable();
                            //cmd.CommandText = "SELECT KeikakuID FROM KeikakuJouhou WHERE KeikakuBangou = '" + planNo + "'";
                            //var sda = new SqlDataAdapter(cmd);
                            //sda.Fill(KeikakuDt);
                            //saibanPlanNo = 0;
                            //if (KeikakuDt != null && KeikakuDt.Rows.Count > 0)
                            //{
                            //    int.TryParse(KeikakuDt.Rows[0][0].ToString(), out saibanPlanNo);
                            //}

                            //計画情報登録更新
                            if (procces.Equals("登録"))
                            {
                                //GlobalMethod.outputLogger(pgmName + methodName, "計画情報の登録", planNo.ToString(), UserInfos[1]);

                                //処理区分が登録のときは新規登録
                                //int saibanPlanNo = GlobalMethod.getSaiban("KeikakuID");
                                saibanPlanNo = GlobalMethod.getSaiban("KeikakuID");


                                cmd.CommandText = InsertSQLText
                                                + saibanPlanNo
                                                + ", N'" + GlobalMethod.ChangeSqlText(fileName, 0, 100) + "' "//ファイル名
                                                + ", N'" + GlobalMethod.ChangeSqlText(Sakuseibi, 0, 0) + "' " //作成日
                                                + ", N'" + GlobalMethod.ChangeSqlText(KashoShibuMei, 0, 50) + "' " //課所支部名
                                                + ", N'" + ChangeSQLString(InputObject[i, 2], 0, 0) + "' " //計画登録日
                                                + ", N'" + ChangeSQLString(InputObject[i, 3], 0, 4) + "' " //売上年度
                                                + ", N'" + ChangeSQLString(planNo, 0, 10) + "' " //計画番号
                                                + ", N'" + ChangeSQLString(InputObject[i, 5], 0, 10) + "' " //前年度計画番号
                                                + ", N'" + ChangeSQLString(InputObject[i, 6], 0, 40) + "' " //前年度案件番号
                                                + ", N'" + ChangeSQLString(InputObject[i, 7], 0, 40) + "' " //前年度受託番号
                                                + ", N'" + ChangeSQLString(InputObject[i, 8], 0, 50) + "' " //前年度発注者名・課名
                                                + ", N'" + ChangeSQLString(InputObject[i, 9], 0, 100) + "' " //前年度業務名称
                                                + ", N'" + DiscriptToValue(ChangeSQLString(InputObject[i, 10], 0, 50), "ShibuMei", "BushoShibuCD", "Mst_Busho") + "' " //部所支部CD
                                                + ", N'" + ChangeSQLString(InputObject[i, 10], 0, 50) + "' " //部所支部名
                                                + ", N'" + DiscriptToValue(ChangeSQLString(InputObject[i, 11], 0, 50), "BushokanriboKamei", "KashoShibuCD", "Mst_Busho") + "' " //課所支部CD
                                                + ", N'" + ChangeSQLString(InputObject[i, 11], 0, 50) + "' " //課所支部名
                                                + ", N'" + DiscriptToValue(ChangeSQLString(InputObject[i, 12], 0, 50), "GyoumuKubun", "GyoumuNarabijunCD", "Mst_GyoumuKubun") + "' " //契約区分CD
                                                + ", N'" + ChangeSQLString(InputObject[i, 12], 0, 50) + "' " //契約区分
                                                + ", N'" + ChangeSQLString(InputObject[i, 13], 0, 100) + "' " //発注者名・課名
                                                + ", N'" + ChangeSQLString(InputObject[i, 14], 0, 150) + "' " //計画案件名
                                                + ", N'" + ChangeSQLInt(InputObject[i, 15], 0, 0) + "' " //調査部配分　落札時見込額（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 16], 0, 0) + "' " //調査部配分　変更見込額（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 17], 0, 0) + "' " //調査部配分　見込額合計（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 18], 0, 0) + "' " //事業普及部配分　落札時見込額（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 19], 0, 0) + "' " //事業普及部配分　変更見込額（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 20], 0, 0) + "' " //事業普及部配分　見込額合計（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 21], 0, 0) + "' " //情報システム部配分　落札時見込額（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 22], 0, 0) + "' " //情報システム部配分　変更見込額（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 23], 0, 0) + "' " //情報システム部配分　見込額合計（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 24], 0, 0) + "' " //総合研究所　落札時見込額（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 25], 0, 0) + "' " //総合研究所　変更見込額（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 26], 0, 0) + "' " //総合研究所　見込額合計（税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 27], 0, 0) + "' " //見込額総合計(税抜）
                                                + ", N'" + ChangeSQLInt(InputObject[i, 28], 0, 0) + "' " //調査業務別　配分　資材調査
                                                + ", N'" + ChangeSQLInt(InputObject[i, 29], 0, 0) + "' " //調査業務別　配分　営繕
                                                + ", N'" + ChangeSQLInt(InputObject[i, 30], 0, 0) + "' " //調査業務別　配分　機器類調査
                                                + ", N'" + ChangeSQLInt(InputObject[i, 31], 0, 0) + "' " //調査業務別　配分　工事費調査
                                                + ", N'" + ChangeSQLInt(InputObject[i, 32], 0, 0) + "' " //調査業務別　配分　産廃調査
                                                + ", N'" + ChangeSQLInt(InputObject[i, 33], 0, 0) + "' " //調査業務別　配分　歩掛調査
                                                + ", N'" + ChangeSQLInt(InputObject[i, 34], 0, 0) + "' " //調査業務別　配分　諸経費調査
                                                + ", N'" + ChangeSQLInt(InputObject[i, 35], 0, 0) + "' " //調査業務別　配分　原価分析調査
                                                + ", N'" + ChangeSQLInt(InputObject[i, 36], 0, 0) + "' " //調査業務別　配分　基準作成改訂
                                                + ", N'" + ChangeSQLInt(InputObject[i, 37], 0, 0) + "' " //調査業務別　配分　公共労費調査
                                                + ", N'" + ChangeSQLInt(InputObject[i, 38], 0, 0) + "' " //調査業務別　配分　労務費公共以外
                                                + ", N'" + ChangeSQLInt(InputObject[i, 39], 0, 0) + "' " //調査業務別　配分　その他調査部
                                                + ", N'" + ChangeSQLInt(InputObject[i, 40], 0, 0) + "' " //調査業務別　配分　合計
                                                + ", " + GetAnkensu(planNo) + " " //案件数
                                                + ", SYSDATETIME()  " //作成日時
                                                + ", N'" + UserInfos[0] + "' " //作成ユーザ
                                                + ", '計画情報取込' " //作成機能
                                                + ", SYSDATETIME()  " //更新日時
                                                + ", N'" + UserInfos[0] + "' " //更新ユーザ
                                                + ", '計画情報取込' " //更新機能
                                                + ", '0' " //削除フラグ
                                                + ")";

                                cmd.ExecuteNonQuery();

                                //GlobalMethod.outputLogger(pgmName + methodName, "変更履歴の登録", planNo.ToString(), UserInfos[1]);
                                // 共通メソッドを使うと、コミットされてしまうので、個別に記載する
                                //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "計画情報を登録しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName, "");
                                cmd.CommandText = Insert_History_SQLText("計画情報を登録しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName);
                                cmd.ExecuteNonQuery();

                            }
                        }

                        //更新処理
                        for (int i = 1; getRowCount - 7 >= i; i++)
                        {

                            //処理区分 部所支部名
                            procces = "";
                            if (InputObject[i, 1] != null)
                            {
                                procces = InputObject[i, 1].ToString();
                            }
                            planNo = "";
                            if (InputObject[i, 4] != null)
                            {
                                planNo = InputObject[i, 4].ToString();
                            }
                            nameBusho = "";
                            if (InputObject[i, 10] != null)
                            {
                                nameBusho = InputObject[i, 10].ToString();
                            }

                            //planNoが取得できない場合、以降空白行とみなし処理終了
                            if (String.IsNullOrEmpty(planNo))
                            {
                                break;
                            }

                            //// 計画IDの取得
                            //var KeikakuDt = new DataTable();
                            //cmd.CommandText = "SELECT KeikakuID FROM KeikakuJouhou WHERE KeikakuBangou = '" + planNo + "'";
                            //var sda = new SqlDataAdapter(cmd);
                            //sda.Fill(KeikakuDt);
                            //saibanPlanNo = 0;
                            //if (KeikakuDt != null && KeikakuDt.Rows.Count > 0)
                            //{
                            //    int.TryParse(KeikakuDt.Rows[0][0].ToString(), out saibanPlanNo);
                            //}

                            //計画情報更新
                            if (procces.Equals("更新"))
                            {
                                // 計画IDの取得
                                var KeikakuDt = new DataTable();
                                cmd.CommandText = "SELECT KeikakuID FROM KeikakuJouhou WHERE KeikakuBangou = '" + planNo + "'";
                                var sda = new SqlDataAdapter(cmd);
                                sda.Fill(KeikakuDt);
                                saibanPlanNo = 0;
                                if (KeikakuDt != null && KeikakuDt.Rows.Count > 0)
                                {
                                    int.TryParse(KeikakuDt.Rows[0][0].ToString(), out saibanPlanNo);
                                }

                                //GlobalMethod.outputLogger(pgmName + methodName, "計画情報の更新", planNo.ToString(), UserInfos[1]);

                                //処理区分が更新のときは更新処理 
                                cmd.CommandText = "UPDATE KeikakuJouhou SET " +
                                    "KeikakuSakuseibi                   = N'" + GlobalMethod.ChangeSqlText(Sakuseibi, 0, 0) + "' " +
                                    ",KeikakuFileMei                    = N'" + GlobalMethod.ChangeSqlText(fileName, 0, 100) + "' " +
                                    ",KeikakuSakuseiKashoShibuMei       = N'" + GlobalMethod.ChangeSqlText(KashoShibuMei, 0, 50) + "' " +
                                    ",KeikakuTourokubi                  = N'" + ChangeSQLString(InputObject[i, 2], 0, 0) + "' " +
                                    ",KeikakuUriageNendo                = N'" + ChangeSQLString(InputObject[i, 3], 0, 4) + "' " +
                                    ",KeikakuZenkaiKeikakuBangou        = N'" + ChangeSQLString(InputObject[i, 5], 0, 10) + "' " +
                                    ",KeikakuZenkaiAnkenBangou          = N'" + ChangeSQLString(InputObject[i, 6], 0, 40) + "' " +
                                    ",KeikakuZenkaiJutakuBangou         = N'" + ChangeSQLString(InputObject[i, 7], 0, 40) + "' " +
                                    ",KeikakuZenkaiHachuushaMeiKaMei    = N'" + ChangeSQLString(InputObject[i, 8], 0, 100) + "' " +
                                    ",KeikakuZenkaiGyoumuMei            = N'" + ChangeSQLString(InputObject[i, 9], 0, 150) + "' " +
                                    ",KeikakuBushoShibuCD               = N'" + DiscriptToValue(ChangeSQLString(InputObject[i, 10], 0, 50), "ShibuMei", "BushoShibuCD", "Mst_Busho") + "' " +
                                    ",KeikakuBushoShibuMei              = N'" + ChangeSQLString(InputObject[i, 10], 0, 50) + "' " +
                                    ",KeikakuKashoShibuCD               = N'" + DiscriptToValue(ChangeSQLString(InputObject[i, 11], 0, 50), "BushokanriboKamei", "KashoShibuCD", "Mst_Busho") + "' " +
                                    ",KeikakuKashoShibuMei              = N'" + ChangeSQLString(InputObject[i, 11], 0, 50) + "' " +
                                    ",KeikakuGyoumuKubun                = N'" + DiscriptToValue(ChangeSQLString(InputObject[i, 12], 0, 50), "GyoumuKubun", "GyoumuNarabijunCD", "Mst_GyoumuKubun") + "' " +
                                    ",KeikakuGyoumuKubunMei             = N'" + ChangeSQLString(InputObject[i, 12], 0, 50) + "' " +
                                    ",KeikakuHachuushaMeiKaMei          = N'" + ChangeSQLString(InputObject[i, 13], 0, 100) + "' " +
                                    ",KeikakuAnkenMei                   = N'" + ChangeSQLString(InputObject[i, 14], 0, 150) + "' " +
                                    ",KeikakuRakusatsumikomigaku        = N'" + ChangeSQLInt(InputObject[i, 15], 0, 0) + "' " +
                                    ",KeikakuHenkoumikomigaku           = N'" + ChangeSQLInt(InputObject[i, 16], 0, 0) + "' " +
                                    ",KeikakuMikomigaku                 = N'" + ChangeSQLInt(InputObject[i, 17], 0, 0) + "' " +
                                    ",KeikakuRakusatsumikomigakuJF      = N'" + ChangeSQLInt(InputObject[i, 18], 0, 0) + "' " +
                                    ",KeikakuHenkoumikomigakuJF         = N'" + ChangeSQLInt(InputObject[i, 19], 0, 0) + "' " +
                                    ",KeikakuMikomigakuJF               = N'" + ChangeSQLInt(InputObject[i, 20], 0, 0) + "' " +
                                    ",KeikakuRakusatsumikomigakuJ       = N'" + ChangeSQLInt(InputObject[i, 21], 0, 0) + "' " +
                                    ",KeikakuHenkoumikomigakuJ          = N'" + ChangeSQLInt(InputObject[i, 22], 0, 0) + "' " +
                                    ",KeikakuMikomigakuJ                = N'" + ChangeSQLInt(InputObject[i, 23], 0, 0) + "' " +
                                    ",KeikakuRakusatsumikomigakuK       = N'" + ChangeSQLInt(InputObject[i, 24], 0, 0) + "' " +
                                    ",KeikakuHenkoumikomigakuK          = N'" + ChangeSQLInt(InputObject[i, 25], 0, 0) + "' " +
                                    ",KeikakuMikomigakuK                = N'" + ChangeSQLInt(InputObject[i, 26], 0, 0) + "' " +
                                    ",KeikakuMikomigakuGoukei           = N'" + ChangeSQLInt(InputObject[i, 27], 0, 0) + "' " +
                                    ",KeikakuShizaiChousa               = N'" + ChangeSQLInt(InputObject[i, 28], 0, 0) + "' " +
                                    ",KeikakuEizen                      = N'" + ChangeSQLInt(InputObject[i, 29], 0, 0) + "' " +
                                    ",KeikakuKikiruiChousa              = N'" + ChangeSQLInt(InputObject[i, 30], 0, 0) + "' " +
                                    ",KeikakuKoujiChousahi              = N'" + ChangeSQLInt(InputObject[i, 31], 0, 0) + "' " +
                                    ",KeikakuSanpaiChousa               = N'" + ChangeSQLInt(InputObject[i, 32], 0, 0) + "' " +
                                    ",KeikakuHokakeChousa               = N'" + ChangeSQLInt(InputObject[i, 33], 0, 0) + "' " +
                                    ",KeikakuShokeihiChousa             = N'" + ChangeSQLInt(InputObject[i, 34], 0, 0) + "' " +
                                    ",KeikakuGenkaBunseki               = N'" + ChangeSQLInt(InputObject[i, 35], 0, 0) + "' " +
                                    ",KeikakuKijunsakusei               = N'" + ChangeSQLInt(InputObject[i, 36], 0, 0) + "' " +
                                    ",KeikakuKoukyouRoumuhi             = N'" + ChangeSQLInt(InputObject[i, 37], 0, 0) + "' " +
                                    ",KeikakuRoumuhiKoukyouigai         = N'" + ChangeSQLInt(InputObject[i, 38], 0, 0) + "' " +
                                    ",KeikakuSonotaChousabu             = N'" + ChangeSQLInt(InputObject[i, 39], 0, 0) + "' " +
                                    ",KeikakuHaibunGoukei               = N'" + ChangeSQLInt(InputObject[i, 40], 0, 0) + "' " +
                                    ",KeikakuAnkensu                    = " + GetAnkensu(planNo) + " " +
                                    ",KeikakuUpdateDate                 = SYSDATETIME() " +
                                    ",KeikakuUpdateUser                 = '" + UserInfos[0] + "' " +
                                    ",KeikakuUpdateProgram              = '計画情報取込' " +
                                    ",KeikakuDeleteFlag                 = '0' " +
                                    "WHERE KeikakuBangou                = '" + planNo + "' ";

                                cmd.ExecuteNonQuery();

                                //GlobalMethod.outputLogger(pgmName + methodName, "変更履歴の登録", planNo.ToString(), UserInfos[1]);
                                // 共通メソッドを使うと、コミットされてしまうので、個別に記載する
                                //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "計画情報を更新しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName, "");
                                cmd.CommandText = Insert_History_SQLText("計画情報を更新しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName);
                                cmd.ExecuteNonQuery();
                            }
                        }

                        //登録更新処理
                        for (int i = 1; getRowCount - 7 >= i; i++)
                        {

                            //処理区分 部所支部名
                            procces = "";
                            if (InputObject[i, 1] != null)
                            {
                                procces = InputObject[i, 1].ToString();
                            }
                            planNo = "";
                            if (InputObject[i, 4] != null)
                            {
                                planNo = InputObject[i, 4].ToString();
                            }
                            nameBusho = "";
                            if (InputObject[i, 10] != null)
                            {
                                nameBusho = InputObject[i, 10].ToString();
                            }

                            //planNoが取得できない場合、以降空白行とみなし処理終了
                            if (String.IsNullOrEmpty(planNo))
                            {
                                break;
                            }

                            //// 計画IDの取得
                            //var KeikakuDt = new DataTable();
                            //cmd.CommandText = "SELECT KeikakuID FROM KeikakuJouhou WHERE KeikakuBangou = '" + planNo + "'";
                            //var sda = new SqlDataAdapter(cmd);
                            //sda.Fill(KeikakuDt);
                            //saibanPlanNo = 0;
                            //if (KeikakuDt != null && KeikakuDt.Rows.Count > 0)
                            //{
                            //    int.TryParse(KeikakuDt.Rows[0][0].ToString(), out saibanPlanNo);
                            //}

                            //計画情報削除
                            if (procces.Equals("削除"))
                            {
                                // 計画IDの取得
                                var KeikakuDt = new DataTable();
                                cmd.CommandText = "SELECT KeikakuID FROM KeikakuJouhou WHERE KeikakuBangou = '" + planNo + "'";
                                var sda = new SqlDataAdapter(cmd);
                                sda.Fill(KeikakuDt);
                                saibanPlanNo = 0;
                                if (KeikakuDt != null && KeikakuDt.Rows.Count > 0)
                                {
                                    int.TryParse(KeikakuDt.Rows[0][0].ToString(), out saibanPlanNo);
                                }

                                //GlobalMethod.outputLogger(pgmName + methodName, "案件情報の更新", planNo.ToString(), UserInfos[1]);

                                // 案件情報から計画情報のクリア
                                cmd.CommandText = "UPDATE AnkenJouhou SET"
                                                + "  AnkenKeikakuBangou = ''"                       // 計画番号
                                                + ", AnkenKeikakuAnkenMei = ''"                     // 計画案件名
                                                + " WHERE AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + planNo + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                //GlobalMethod.outputLogger(pgmName + methodName, "計画情報の削除", planNo.ToString(), UserInfos[1]);

                                // 計画情報の削除
                                cmd.CommandText = "DELETE FROM KeikakuJouhou"
                                                + " WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + planNo + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                //GlobalMethod.outputLogger(pgmName + methodName, "変更履歴の登録", planNo.ToString(), UserInfos[1]);
                                // 共通メソッドを使うと、コミットされてしまうので、個別に記載する
                                //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "計画情報を削除しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName, "");
                                cmd.CommandText = Insert_History_SQLText("計画情報を削除しました ID:" + saibanPlanNo + " 計画番号:" + ChangeSQLString(planNo, 0, 10), pgmName + methodName);
                                cmd.ExecuteNonQuery();

                            }
                        }

                        transaction.Commit();
                    }
                    catch (Exception e)
                    {
                        transaction.Rollback();
                        GlobalMethod.outputLogger(pgmName + methodName, "計画情報取込エラー1", e.ToString(), UserInfos[1]);
                        throw;
                    }
                    finally
                    {
                        // Rangeもプロセス開放しないといけない
                        Marshal.ReleaseComObject(w_rgnName);
                    }
                    conn.Close();
                }
                return true;
            }
            catch (Exception e)
            {
                GlobalMethod.outputLogger(pgmName + methodName, "計画情報取込エラー2", e.ToString(), UserInfos[1]);
                return false;
            }
        }

        private string DiscriptToValue(string Discript, string DiscriptCol, string ValueCol, string Table)
        {
            string value = "";
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                //エラーメッセージ
                conn.Open();
                var cmd = conn.CreateCommand();
                var dt = new DataTable();
                //SQL生成
                cmd.CommandText = "SELECT TOP 1 " + ValueCol + " " +
                "FROM " + Table + " " +
                "WHERE " + DiscriptCol + " COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(Discript, 0, 0) + "' ";
                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
                if (dt.Rows.Count >= 1)
                {
                    DataRow dr = dt.Rows[0];
                    value = dr[ValueCol].ToString();
                }

                conn.Close();
            }

            return value;
        }
        private int GetAnkensu(string No)
        {
            int ankensu = 0;
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                //エラーメッセージ
                conn.Open();
                var cmd = conn.CreateCommand();
                var dt = new DataTable();
                //SQL生成
                cmd.CommandText = "SELECT  COUNT(AnkenJouhouID) AS 'Ankensu' " +
                "FROM AnkenJouhou " +
                "WHERE AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(No, 0, 0) + "' " +
                "AND AnkenDeleteFlag = 0";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
                if (dt.Rows.Count >= 1)
                {
                    ankensu = int.Parse(dt.Rows[0][0].ToString());
                }

                conn.Close();
            }

            return ankensu;
        }

        // 検索ボタン
        private void button2_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            ErrorMessage.Text = "";

            if (src_4.Text == "")
            {
                // 全国検索の為データ件数が多くなりますが、よろしいですか？ 20210225 聞かない
                //if (GlobalMethod.outputMessage("I10501", "") == DialogResult.OK)
                //{
                    get_date();
                //}
            }
            else
            {
                get_date();
            }

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ErrorMessage.Text = "";
            Popup_FileError form = new Popup_FileError(errorUser, readCount);
            form.ShowDialog();
        }

        private void get_date()
        {
            //String nendo1 = src_1.Text.Substring(0, 4);
            //String nendo2 = nendo1;
            //if (src_3.Checked)
            //{
            //    nendo2 = (int.Parse(nendo1) - 2).ToString();
            //}

            String nendo1 = "";
            String nendo2 = "";

            int year = 0;

            // 年度計算
            // 売上年度
            if (src_1.Text != "")
            {
                nendo1 = src_1.Text.Substring(0, 4);
                nendo2 = nendo1;
                if (src_3.Checked)
                {
                    nendo2 = (int.Parse(nendo1) - 2).ToString();
                }
            }
            else
            {
                year = int.Parse(GlobalMethod.GetTodayNendo());
                nendo1 = year.ToString();
                year = year - 4;
                nendo2 = year.ToString();
            }

            int FromNendo;
            if (!int.TryParse(src_1.SelectedValue.ToString(), out FromNendo))
            {
                FromNendo = DateTime.Today.Year;
            }
            int ToNendo = FromNendo + 1;

            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            var dt2 = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = @"SELECT KeikakuID,KeikakuUriageNendo,KeikakuBangou,KeikakuBushoShibuCD
                    ,KeikakuZenkaiKeikakuBangou,KeikakuZenkaiAnkenBangou,KeikakuZenkaiJutakuBangou
                    ,KeikakuHachuushaMeiKaMei,KeikakuAnkenMei,KeikakuGyoumuKubun,KeikakuTourokubi
                    ,KeikakuKashoShibuCD
                    ,KeikakuRakusatsumikomigaku,KeikakuHenkoumikomigaku,KeikakuMikomigaku
                    ,KeikakuRakusatsumikomigakuJF,KeikakuHenkoumikomigakuJF,KeikakuMikomigakuJF
                    ,KeikakuRakusatsumikomigakuJ,KeikakuHenkoumikomigakuJ,KeikakuMikomigakuJ
                    ,KeikakuRakusatsumikomigakuK,KeikakuHenkoumikomigakuK,KeikakuMikomigakuK
                    ,KeikakuMikomigakuGoukei,KeikakuAnkensu 
                    FROM KeikakuJouhou " +
                    "LEFT JOIN Mst_Busho ON KeikakuBushoShibuCD = BushoShibuCD AND KeikakuKashoShibuCD = KashoShibuCD " +
                    //"AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                    //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                    "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' ) " +
                    "WHERE KeikakuUriageNendo IS NOT NULL and KeikakuUriageNendo <= '" + nendo1 + "' and KeikakuUriageNendo >= '" + nendo2 + "' ";
                if (src_4.Text != "")
                {
                    if (src_4.SelectedValue.ToString() == "127100")
                    {
                        cmd.CommandText += "  and ( GyoumuBushoCD like '1271%' or GyoumuBushoCD = '127220' ) ";
                    }
                    else
                    {
                        cmd.CommandText += "  and GyoumuBushoCD = '" + src_4.SelectedValue + "'";
                    }
                }
                if (src_5.Text != "")
                {
                    cmd.CommandText += "  and KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_5.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                if (src_6.Text != "")
                {
                    cmd.CommandText += "  and KeikakuHachuushaMeiKaMei COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_6.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                if (src_7.Text != "")
                {
                    cmd.CommandText += "  and KeikakuAnkenMei COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_7.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                if (src_8.Text != "")
                {
                    cmd.CommandText += "  and KeikakuGyoumuKubun = N'" + src_8.SelectedValue + "' ";
                }
                if (src_9.Text != "")
                {
                    cmd.CommandText += "  and KeikakuZenkaiAnkenBangou COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_9.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                if (src_10.Text != "")
                {
                    cmd.CommandText += "  and KeikakuZenkaiJutakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_10.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                if (src_11.Text != "")
                {
                    if (src_12.Text == "=")
                    {
                        cmd.CommandText += "  and KeikakuAnkensu = " + src_11.Text + " ";
                    }
                    else
                    {
                        cmd.CommandText += "  and KeikakuAnkensu >= " + src_11.Text + " ";
                    }
                }
                cmd.CommandText += " and KeikakuBangou not like '%-P9%' ";
                cmd.CommandText += " ORDER BY KeikakuBangou DESC ";
                var sda = new SqlDataAdapter(cmd);

                ListData.Clear();
                sda.Fill(ListData);
            }
            Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / int.Parse(src_13.Text))).ToString();
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

        private void set_data(int pagenum)
        {
            c1FlexGrid1.Rows.Count = 2;
            c1FlexGrid1.AllowAddNew = true;
            int startrow = (pagenum - 1) * int.Parse(src_13.Text);
            int addnum = ListData.Rows.Count - startrow;
            if (addnum > int.Parse(src_13.Text))
            {
                addnum = int.Parse(src_13.Text);
            }
            for (int r = 0; r < addnum; r++)
            {
                c1FlexGrid1.Rows.Add();
                c1FlexGrid1[r + 2, 2] = ListData.Rows[startrow + r][0];
                c1FlexGrid1[r + 2, 3] = ListData.Rows[startrow + r][1];
                c1FlexGrid1[r + 2, 4] = ListData.Rows[startrow + r][2];
                c1FlexGrid1[r + 2, 5] = ListData.Rows[startrow + r][3];
                c1FlexGrid1[r + 2, 6] = ListData.Rows[startrow + r][4];
                c1FlexGrid1[r + 2, 7] = ListData.Rows[startrow + r][5];
                c1FlexGrid1[r + 2, 8] = ListData.Rows[startrow + r][6];
                c1FlexGrid1[r + 2, 9] = ListData.Rows[startrow + r][7];
                c1FlexGrid1[r + 2, 10] = ListData.Rows[startrow + r][8];
                c1FlexGrid1[r + 2, 11] = ListData.Rows[startrow + r][9];
                c1FlexGrid1[r + 2, 12] = ListData.Rows[startrow + r][10];
                c1FlexGrid1[r + 2, 13] = ListData.Rows[startrow + r][11];
                c1FlexGrid1[r + 2, 14] = ListData.Rows[startrow + r][12];
                c1FlexGrid1[r + 2, 15] = ListData.Rows[startrow + r][13];
                c1FlexGrid1[r + 2, 16] = ListData.Rows[startrow + r][14];
                c1FlexGrid1[r + 2, 17] = ListData.Rows[startrow + r][15];
                c1FlexGrid1[r + 2, 18] = ListData.Rows[startrow + r][16];
                c1FlexGrid1[r + 2, 19] = ListData.Rows[startrow + r][17];
                c1FlexGrid1[r + 2, 20] = ListData.Rows[startrow + r][18];
                c1FlexGrid1[r + 2, 21] = ListData.Rows[startrow + r][19];
                c1FlexGrid1[r + 2, 22] = ListData.Rows[startrow + r][20];
                c1FlexGrid1[r + 2, 23] = ListData.Rows[startrow + r][21];
                c1FlexGrid1[r + 2, 24] = ListData.Rows[startrow + r][22];
                c1FlexGrid1[r + 2, 25] = ListData.Rows[startrow + r][23];
                c1FlexGrid1[r + 2, 26] = ListData.Rows[startrow + r][24];
                c1FlexGrid1[r + 2, 27] = ListData.Rows[startrow + r][25];
                c1FlexGrid1.Rows[r + 2].Height = 40;
            }
            c1FlexGrid1.AllowAddNew = false;
            if (c1FlexGrid1.Rows.Count > 2)
            {
                C1.Win.C1FlexGrid.CellRange cr;
                cr = c1FlexGrid1.GetCellRange(2, 1, c1FlexGrid1.Rows.Count - 1, 1);
                cr.Image = Image.FromFile("Resource/Image/file_presentation1.png");
                c1FlexGrid1.Select(2, 2, true);
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

        private void button4_Click(object sender, EventArgs e)
        {
            Entry_keikaku_Search form = new Entry_keikaku_Search();
            form.UserInfos = this.UserInfos;
            form.Show();
            this.Close();
        }

        // 計画情報一覧出力ボタン押下
        private void button1_Click(object sender, EventArgs e)
        {
            Popup_Loading Loading = new Popup_Loading();
            Loading.StartPosition = FormStartPosition.CenterScreen;
            Loading.Show();

            // COMExceptionが発生したかのフラグ false:正常 true:エラー
            comexceptionFlg = false;
            comexceptionCnt = 0;

            // エクセル出力 + やり直し3回
            for(int i = 0; i <= 3; i++)
            {
                excelExport();
                // 出力に成功したら終わり
                if (comexceptionFlg == false)
                {
                    break;
                }
            }
            Loading.Close();
        }

        // エクセル出力
        private void excelExport()
        {
            comexceptionFlg = false;
            try
            {
                Application ExcelApp = new Application();
                Workbook wb = null;
                Worksheet ws = new Worksheet();
                ErrorMessage.Text = "";

                string PrintDataPatternName = "";
                string PrintName = "";
                string PrintFileName = "";
                string PrintHistoryDownLoadFileName = "";
                string BushokanriboKamei = "";
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();

                    // プリントマスタの取得
                    cmd.CommandText = "SELECT"
                                    + " PrintListID"
                                    + ", PrintDataPatternName"
                                    + ", PrintName"
                                    + ", PrintFileName"
                                    + ", PrintDownloadFileName"
                                    + " FROM Mst_PrintList"
                                    + " LEFT JOIN M_PrintPattern ON Mst_PrintList.PrintDataPattern = M_PrintPattern.PrintDataPattern"
                                    + " WHERE Mst_PrintList.PrintDataPattern = 79 AND Mst_PrintList.PrintDelFlg = 0"
                                    ;

                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    DataTable dtPlintList = new DataTable();
                    sda.Fill(dtPlintList);
                    if (dtPlintList.Rows.Count > 0)
                    {
                        PrintDataPatternName = dtPlintList.Rows[0][1].ToString();
                        PrintName = dtPlintList.Rows[0][2].ToString();
                        PrintFileName = dtPlintList.Rows[0][3].ToString();
                        PrintHistoryDownLoadFileName = dtPlintList.Rows[0][4].ToString();
                    }

                    // 課名の取得
                    cmd.CommandText = "SELECT"
                                    + " BushokanriboKamei"
                                    + " FROM Mst_Busho"
                                    + " WHERE GyoumuBushoCD = '" + UserInfos[2] + "'"
                                    ;
                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    DataTable Mst_Busho = new DataTable();
                    sda.Fill(Mst_Busho);
                    if (Mst_Busho.Rows.Count > 0)
                    {
                        if (Mst_Busho.Rows[0][0] != null)
                        {
                            BushokanriboKamei = Mst_Busho.Rows[0][0].ToString();
                        }
                    }

                    conn.Close();
                }

                string plusName = "_" + src_1.SelectedValue.ToString();
                if (src_4.Text == "")
                {
                    plusName += "_部所支部名";
                }
                else
                {
                    plusName += "_" + src_4.Text;
                }

                try
                {
                    //Excel取込
                    ExcelApp.DisplayAlerts = false;
                    string[] fileName = GlobalMethod.GetHinagataPath(700);
                    wb = getExcelFile(fileName[0], ExcelApp);

                    if (wb == null)
                    {
                        return;
                    }

                    //シート取込　testファイル用出力データ取得・セット処理
                    setExcelFile(wb, ws);
                    //string ExcelPath = "Work";
                    long longWorkRenban = GlobalMethod.getSaiban("WorkCopyRenban");
                    string ExcelPath = GlobalMethod.GetCommonValue1("WORK_FOLDER") + longWorkRenban.ToString().PadLeft(14, '0');    // CommonMasterからWORK_FOLDERを取得して連番フォルダを用意する
                    string ExcelName = wb.Name.Split('.')[0] + plusName + "." + wb.Name.Split('.')[1];  // 拡張子とファイル名を分割

                    string path = System.IO.Path.Combine(new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).Directory.FullName, ExcelPath);
                    if (!File.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    try
                    {
                        //GlobalMethod.Get_WorkFolder();
                        wb.SaveAs(System.IO.Path.Combine(new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).Directory.FullName, ExcelPath, ExcelName));

                        // 印刷履歴の出力
                        string[] report_data = new string[11] { "", "", "", "", "", "", "", "", "", "", "" };

                        // 0.部所CD
                        report_data[0] = UserInfos[2];
                        // 1.部所名
                        //report_data[1] = UserInfos[3];
                        report_data[1] = BushokanriboKamei;
                        // 2.個人CD
                        report_data[2] = UserInfos[0];
                        // 3.職員名
                        report_data[3] = UserInfos[1];
                        // 4.業務名
                        report_data[4] = "エントリくん";
                        // 5.特調奉行名
                        report_data[5] = "";
                        // 6.画面・機能名
                        report_data[6] = "計画情報検索画面";
                        // 7.帳票分類名
                        report_data[7] = PrintDataPatternName;
                        // 8.帳票名
                        report_data[8] = PrintName;
                        // 9.雛型ファイル名
                        report_data[9] = PrintFileName;
                        // 10.ダウンロードファイル名
                        //report_data[10] = PrintHistoryDownLoadFileName;
                        report_data[10] = ExcelName;

                        GlobalMethod.Insert_PrintHistory(report_data);
                    }
                    catch (Exception)
                    {
                        ErrorMessage.Text = "EXCELファイルの作成に失敗いたしました。";
                        ErrorMessage.Visible = true;
                    }

                    if (ErrorMessage.Text == "")
                    {
                        Popup_Download form = new Popup_Download();
                        form.TopLevel = false;
                        this.Controls.Add(form);
                        //form.ExcelPath = ExcelPath;
                        form.ExcelPath = System.IO.Path.Combine(ExcelPath, ExcelName);
                        form.ExcelName = ExcelName;
                        form.Dock = DockStyle.Bottom;

                        form.Show();
                        form.BringToFront();
                    }


                    wb.Close(false);
                    ExcelApp.Quit();
                }
                finally
                {
                    //Loading.Close();
                    try
                    {
                        //Excelのオブジェクトを開放し忘れているとプロセスが落ちないため注意
                        Marshal.ReleaseComObject(ws);
                        if (wb != null)
                        {
                            Marshal.ReleaseComObject(wb);
                        }
                        Marshal.ReleaseComObject(ExcelApp);
                    }
                    catch (Exception)
                    {
                    }
                    ws = null;
                    wb = null;
                    ExcelApp = null;
                    GC.Collect();
                }
                // バックグラウンドとなっているExcelプロセスをKILL
                excelProcessKill();
            }
            catch (System.Runtime.InteropServices.COMException come)
            {
                // COMExceptionが発生したかのフラグ false:正常 true:エラー
                comexceptionFlg = true;
                comexceptionCnt += 1;
            }
        }

        private void setExcelFile(Workbook wb, Worksheet ws)
        {
            //Excelシートのインスタンスを作る
            //インデックスは1始まり
            int u = wb.Sheets.Count;
            dynamic xlSheet = null;
            xlSheet = wb.Sheets[u];
            Excel.Range w_rgn = xlSheet.Cells;
            try
            {
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();

                    int FromNendo;
                    if (!int.TryParse(src_1.SelectedValue.ToString(), out FromNendo))
                    {
                        FromNendo = DateTime.Today.Year;
                    }
                    int ToNendo = FromNendo + 1;

                    //部所支部データ取得
                    var dt = new DataTable();
                    //SQL生成
                    cmd.CommandText = "SELECT DISTINCT " +
                    "BushoShibuCD ,ShibuMei " +
                    "FROM " + "Mst_Busho " +
                    "WHERE ISNULL(BushoShibuCD,'') <> '' AND GyoumuBushoCD < '999990' and BushoKeikakuHyoujiFlg = '1' " +
                "AND NOT GyoumuBushoCD LIKE '121%' " +
                //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' " +
                "AND BushoNewOld <= 1 AND ISNULL(BushoDeleteFlag,0) = 0 AND (BushoShibuCD IS NOT NULL OR KashoShibuCD IS NOT NULL) " +
                //"AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
                    "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' )";

                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);

                    //データセット
                    String[,] mData1 = new string[dt.Rows.Count, dt.Columns.Count];
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int k = 0; k < dt.Columns.Count; k++)
                        {
                            mData1[i, k] = dt.Rows[i][k].ToString();
                        }
                    }

                    w_rgn = xlSheet.Range("A3:B" + (dt.Rows.Count + 2));
                    w_rgn.Value = mData1;


                    //課所支部データ取得
                    var dt2 = new DataTable();
                    //SQL生成
                    cmd.CommandText = "SELECT DISTINCT " +
                    "KashoShibuCD,BushokanriboKamei " +
                    "FROM " + "Mst_Busho " +
                    "WHERE ISNULL(KashoShibuCD,'') <> '' AND GyoumuBushoCD < '999990' and BushoKeikakuHyoujiFlg = '1' " +
                "AND NOT GyoumuBushoCD LIKE '121%' " +
                //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' " +
                //"AND BushoNewOld <= 1 AND ISNULL(BushoDeleteFlag,0) = 0 AND (BushoShibuCD IS NOT NULL OR KashoShibuCD IS NOT NULL) " +
                //"AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
                    "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' )";

                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt2);

                    String[,] mData2 = new string[dt2.Rows.Count, dt2.Columns.Count];
                    for (int i = 0; i < dt2.Rows.Count; i++)
                    {
                        for (int k = 0; k < dt2.Columns.Count; k++)
                        {
                            mData2[i, k] = dt2.Rows[i][k].ToString();
                        }
                    }

                    w_rgn = xlSheet.Range("D3:E" + (dt2.Rows.Count + 2));
                    w_rgn.Value = mData2;

                    //契約区分データ取得
                    var dt3 = new DataTable();
                    //SQL生成
                    cmd.CommandText = "SELECT " +
                    "GyoumuNarabijunCD ,GyoumuKubun " +
                    "FROM " + "Mst_GyoumuKubun " +
                    "WHERE GyoumuNarabijunCD < 100 ";

                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt3);

                    //データセット
                    String[,] mData3 = new string[dt3.Rows.Count, dt3.Columns.Count];
                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        for (int k = 0; k < dt3.Columns.Count; k++)
                        {
                            mData3[i, k] = dt3.Rows[i][k].ToString();
                        }
                    }

                    w_rgn = xlSheet.Range("G3:H" + (dt3.Rows.Count + 2));
                    w_rgn.Value = mData3;


                    //計画データ
                    //ws = wb.Sheets[u - 1];
                    ws = wb.Sheets["取り込みシート"];
                    ws.Select(Type.Missing);
                    w_rgn = ws.Cells;
                    var dt4 = new DataTable();

                    w_rgn[2, 3].Value = DateTime.Today;

                    //SQL生成
                    String nendo1 = src_1.Text.Substring(0, 4);
                    String nendo2 = nendo1;
                    if (src_3.Checked)
                    {
                        nendo2 = (int.Parse(nendo1) - 2).ToString();
                    }
                    cmd.CommandText = "SELECT"
                                    + " '更新'"                                                                 // 0.処理区分
                                    + ", CONVERT(NVARCHAR, KeikakuTourokubi, 111) AS 'KeikakuTourokubi'"        // 1.計画登録日
                                    + ", KeikakuUriageNendo"                                                    // 2.売上年度
                                    + ", KeikakuBangou"                                                         // 3.計画番号
                                    + ", KeikakuZenkaiKeikakuBangou"                                            // 4.前年度計画番号
                                    + ", KeikakuZenkaiAnkenBangou"                                              // 5.前年度案件番号
                                    + ", KeikakuZenkaiJutakuBangou"                                             // 6.前年度受託番号
                                    + ", KeikakuZenkaiHachuushaMeiKaMei"                                        // 7.前年度発注者名・課名
                                    + ", KeikakuZenkaiGyoumuMei"                                                // 8.前年度業務名称
                                    + ", KeikakuBushoShibuMei"                                                  // 9.部所支部名
                                    + ", KeikakuKashoShibuMei"                                                  // 10.課所支部名
                                    + ", KeikakuGyoumuKubunMei"                                                 // 11.契約区分
                                    + ", KeikakuHachuushaMeiKaMei"                                              // 12.発注者名・課名
                                    + ", KeikakuAnkenMei"                                                       // 13.計画案件名
                                    + ", KeikakuRakusatsumikomigaku"                                            // 14.調査部配分 落札見込額（税抜）
                                    + ", KeikakuHenkoumikomigaku"                                               // 15.調査部配分 変更見込額（税抜） 
                                    + ", KeikakuMikomigaku"                                                     // 16.調査部配分 見込額合計（税抜）           ※出力対象外
                                    + ", KeikakuRakusatsumikomigakuJF"                                          // 17.事業普及部配分 落札見込額（税抜）
                                    + ", KeikakuHenkoumikomigakuJF"                                             // 18.事業普及部配分 変更見込額（税抜）
                                    + ", KeikakuMikomigakuJF"                                                   // 19.事業普及部配分 見込額合計（税抜）       ※出力対象外
                                    + ", KeikakuRakusatsumikomigakuJ"                                           // 20.情報システム部配分 落札見込額（税抜）
                                    + ", KeikakuHenkoumikomigakuJ"                                              // 21.情報システム部配分 変更見込額（税抜）
                                    + ", KeikakuMikomigakuJ"                                                    // 22.情報システム部配分 見込額合計（税抜）   ※出力対象外
                                    + ", KeikakuRakusatsumikomigakuK"                                           // 23.総合研究所配分 落札見込額（税抜）
                                    + ", KeikakuHenkoumikomigakuK"                                              // 24.総合研究所配分 変更見込額（税抜）
                                    + ", KeikakuMikomigakuK"                                                    // 25.総合研究所配分 見込額合計（税抜）       ※出力対象外
                                    + ", KeikakuMikomigakuGoukei"                                               // 26.見込額総合計（税抜）                    ※出力対象外
                                    + ", KeikakuShizaiChousa"                                                   // 27.調査業務別 配分 資材調査
                                    + ", KeikakuEizen"                                                          // 28.調査業務別 配分 営繕
                                    + ", KeikakuKikiruiChousa"                                                  // 29.調査業務別 配分 機器類調査
                                    + ", KeikakuKoujiChousahi"                                                  // 30.調査業務別 配分 工事費調査
                                    + ", KeikakuSanpaiChousa"                                                   // 31.調査業務別 配分 産廃調査
                                    + ", KeikakuHokakeChousa"                                                   // 32.調査業務別 配分 歩掛調査
                                    + ", KeikakuShokeihiChousa"                                                 // 33.調査業務別 配分 諸経費調査
                                    + ", KeikakuGenkaBunseki"                                                   // 34.調査業務別 配分 原価分析調査
                                    + ", KeikakuKijunsakusei"                                                   // 35.調査業務別 配分 基準作成改訂
                                    + ", KeikakuKoukyouRoumuhi"                                                 // 36.調査業務別 配分 公共労費調査
                                    + ", KeikakuRoumuhiKoukyouigai"                                             // 37.調査業務別 配分 労務費公共以外
                                    + ", KeikakuSonotaChousabu"                                                 // 38.調査業務別 配分 その他調査部
                                    + ", KeikakuHaibunGoukei"                                                   // 39.調査業務別 配分 合計                    ※出力対象外
                                    + " FROM KeikakuJouhou"
                                    + " LEFT JOIN Mst_Busho"
                                    + " ON KeikakuBushoShibuCD = BushoShibuCD"
                                    + " AND KeikakuKashoShibuCD = KashoShibuCD"
                                    //+ " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) "
                                    //+ " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) "
                                    + " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) "
                                    + " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' ) "
                                    + " WHERE KeikakuUriageNendo IS NOT NULL"
                                    + " and KeikakuUriageNendo <= '" + nendo1 + "'"
                                    + " and KeikakuUriageNendo >= '" + nendo2 + "'"
                                    ;
                    if (src_4.Text != "")
                    {
                        if (src_4.SelectedValue.ToString() == "127100")
                        {
                            cmd.CommandText += "  and ( GyoumuBushoCD like '1271%' or GyoumuBushoCD = '127220' ) ";
                        }
                        else
                        {
                            cmd.CommandText += "  and GyoumuBushoCD = '" + src_4.SelectedValue + "'";
                        }
                    }
                    cmd.CommandText += " ORDER BY KeikakuBangou DESC ";

                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt4);
                    conn.Close();


                    //xlSheet = wb.Sheets[u - 1];
                    xlSheet = wb.Sheets["取り込みシート"];

                    // 出力対象のセルをレンジでまとめるため、配列を複数用意する。
                    String[,] pData1 = new string[dt4.Rows.Count, 14];      // B～O      処理区分 ～ 計画案件名
                    int[,] pData2 = new int[dt4.Rows.Count, 2];             // P～Q      調査部配分 落札時見込額（税抜） ～ 変更見込額（税抜）
                    int[,] pData3 = new int[dt4.Rows.Count, 2];             // S～T      事業普及部配分 落札時見込額（税抜） ～ 変更見込額（税抜）
                    int[,] pData4 = new int[dt4.Rows.Count, 2];             // V～W      情報システム部配分 落札時見込額（税抜） ～ 変更見込額（税抜）
                    int[,] pData5 = new int[dt4.Rows.Count, 2];             // Y～Z      総合研究所 落札時見込額（税抜） ～ 変更見込額（税抜）
                    int[,] pData6 = new int[dt4.Rows.Count, 12];            // AC～AN    調査業務別　配分 資材調査 ～ その他調査部

                    int colCount_moto = 0;  // 編集元配列のカウント（編集先が変わっても連続する）
                    int colCount_saki = 0;  // 編集先配列のカウント（編集先が変わるごとリセット）
                    for (int i = 0; i < dt4.Rows.Count; i++)
                    {
                        //for (int k = 0; k < 14; k++)
                        //{
                        //    pData1[i, k] = dt4.Rows[i][k].ToString();
                        //}
                        //for (int k = 15; k < 17; k++)
                        //{
                        //    pData2[i, k - 15] = int.Parse(((decimal)dt4.Rows[i][k]).ToString("F0"));
                        //}
                        //for (int k = 17; k < 19; k++)
                        //{
                        //    pData3[i, k - 17] = int.Parse(((decimal)dt4.Rows[i][k]).ToString("F0"));
                        //}
                        //for (int k = 20; k < 22; k++)
                        //{
                        //    pData4[i, k - 20] = int.Parse(((decimal)dt4.Rows[i][k]).ToString("F0"));
                        //}
                        //for (int k = 23; k < 25; k++)
                        //{
                        //    pData5[i, k - 23] = int.Parse(((decimal)dt4.Rows[i][k]).ToString("F0"));
                        //}
                        //for (int k = 27; k < 39; k++)
                        //{
                        //    pData6[i, k - 27] = int.Parse(((decimal)dt4.Rows[i][k]).ToString("F0"));
                        //}

                        colCount_moto = 0;
                        colCount_saki = 0;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 0.処理区分
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 1.計画登録日
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 2.売上年度
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 3.計画番号
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 4.前年度計画番号
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 5.前年度案件番号
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 6.前年度受託番号
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 7.前年度発注者名・課名
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 8.前年度業務名称
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 9.部所支部名
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 10.課所支部名
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 11.契約区分
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 12.発注者名・課名
                        colCount_saki += 1;
                        pData1[i, colCount_saki] = dt4.Rows[i][colCount_moto + colCount_saki].ToString();     // 13.計画案件名
                        colCount_saki += 1;

                        colCount_moto += colCount_saki;
                        colCount_saki = 0;
                        pData2[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 14.調査部配分 落札見込額（税抜）
                        colCount_saki += 1;
                        pData2[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 15.調査部配分 変更見込額（税抜） 
                        colCount_saki += 1;
                        // 16.調査部配分 見込額合計（税抜）           ※出力対象外 
                        colCount_saki += 1; 

                        colCount_moto += colCount_saki;
                        colCount_saki = 0;
                        pData3[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 17.事業普及部配分 落札見込額（税抜）
                        colCount_saki += 1;
                        pData3[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 18.事業普及部配分 変更見込額（税抜）
                        colCount_saki += 1;
                        // 19.事業普及部配分 見込額合計（税抜）       ※出力対象外
                        colCount_saki += 1;

                        colCount_moto += colCount_saki;
                        colCount_saki = 0;
                        pData4[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 20.情報システム部配分 落札見込額（税抜）
                        colCount_saki += 1;
                        pData4[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 21.情報システム部配分 変更見込額（税抜）
                        colCount_saki += 1;
                        // 22.情報システム部配分 見込額合計（税抜）   ※出力対象外
                        colCount_saki += 1;

                        colCount_moto += colCount_saki;
                        colCount_saki = 0;
                        pData5[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 23.総合研究所配分 落札見込額（税抜）
                        colCount_saki += 1;
                        pData5[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 24.総合研究所配分 変更見込額（税抜）
                        colCount_saki += 1;
                        // 25.総合研究所配分 見込額合計（税抜）       ※出力対象外
                        colCount_saki += 1;
                        // 26.見込額総合計（税抜）                    ※出力対象外
                        colCount_saki += 1;

                        colCount_moto += colCount_saki;
                        colCount_saki = 0;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 27.調査業務別 配分 資材調査
                        colCount_saki += 1;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 28.調査業務別 配分 営繕
                        colCount_saki += 1;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 29.調査業務別 配分 機器類調査
                        colCount_saki += 1;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 30.調査業務別 配分 工事費調査
                        colCount_saki += 1;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 31.調査業務別 配分 産廃調査
                        colCount_saki += 1;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 32.調査業務別 配分 歩掛調査
                        colCount_saki += 1;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 33.調査業務別 配分 諸経費調査
                        colCount_saki += 1;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 34.調査業務別 配分 原価分析調査
                        colCount_saki += 1;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 35.調査業務別 配分 基準作成改訂
                        colCount_saki += 1;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 36.調査業務別 配分 公共労費調査
                        colCount_saki += 1;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 37.調査業務別 配分 労務費公共以外
                        colCount_saki += 1;
                        pData6[i, colCount_saki] = int.Parse(((decimal)dt4.Rows[i][colCount_moto + colCount_saki]).ToString("F0"));      // 38.調査業務別 配分 その他調査部
                        colCount_saki += 1;
                        // 39.調査業務別 配分 合計                    ※出力対象外
                        colCount_saki += 1;
                    }

                    //w_rgn = xlSheet.Range("B8:O" + (dt4.Rows.Count + 7));
                    //w_rgn.Value = pData1;

                    //w_rgn = xlSheet.Range("P8:Q" + (dt4.Rows.Count + 7));
                    //w_rgn.Value = pData2;

                    //w_rgn = xlSheet.Range("S8:T" + (dt4.Rows.Count + 7));
                    //w_rgn.Value = pData3;

                    //w_rgn = xlSheet.Range("V8:W" + (dt4.Rows.Count + 7));
                    //w_rgn.Value = pData4;

                    //w_rgn = xlSheet.Range("Y8:Z" + (dt4.Rows.Count + 7));
                    //w_rgn.Value = pData5;

                    //w_rgn = xlSheet.Range("AC8:AN" + (dt4.Rows.Count + 7));
                    //w_rgn.Value = pData6;

                    // 開始行、終了行の指定
                    int startRow = 8;
                    int endRow = dt4.Rows.Count + 7;

                    // 出力対象のセルを連続するセルでまとめて、配列ごと書き出しを行う。
                    w_rgn = xlSheet.Range("B" + startRow + ":" + "O" + endRow);
                    w_rgn.Value = pData1;

                    w_rgn = xlSheet.Range("P" + startRow + ":" + "Q" + endRow);
                    w_rgn.Value = pData2;

                    w_rgn = xlSheet.Range("S" + startRow + ":" + "T" + endRow);
                    w_rgn.Value = pData3;

                    w_rgn = xlSheet.Range("V" + startRow + ":" + "W" + endRow);
                    w_rgn.Value = pData4;

                    w_rgn = xlSheet.Range("Y" + startRow + ":" + "Z" + endRow);
                    w_rgn.Value = pData5;

                    w_rgn = xlSheet.Range("AC" + startRow + ":" + "AN" + endRow);
                    w_rgn.Value = pData6;

                }
            }
            finally
            {
                Marshal.ReleaseComObject(w_rgn);
                Marshal.ReleaseComObject(xlSheet);
                xlSheet = null;
                GC.Collect();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            ErrorMessage.Text = "";
            set_defalt();

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            ErrorMessage.Text = "";

            Paging_now.Text = (1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();

        }
        private void pictureBox4_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            ErrorMessage.Text = "";

            Paging_now.Text = (int.Parse(Paging_now.Text) + 1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            ErrorMessage.Text = "";

            Paging_now.Text = (int.Parse(Paging_now.Text) - 1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            ErrorMessage.Text = "";

            Paging_now.Text = (int.Parse(Paging_all.Text)).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
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
        }

        private void src_3_Click(object sender, EventArgs e)
        {
            set_combo_shibu(src_1.SelectedValue.ToString());
        }
        private void textbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void c1FlexGrid1_BeforeSort(object sender, C1.Win.C1FlexGrid.SortColEventArgs e)
        {
            DataView dv = new DataView(ListData);
            dv.Sort = ListData.Columns[e.Col - 2].ColumnName;
            if (c1FlexGrid1.Cols[e.Col].Sort == C1.Win.C1FlexGrid.SortFlags.Ascending)
            {
                dv.Sort += " DESC";
            }
            ListData = dv.ToTable();
            set_data(int.Parse(Paging_now.Text));
        }
        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        // バックグラウンドとなっているExcleのプロセスをKillする
        private void excelProcessKill()
        {
            // 他のExcelを起動していてもバックグラウンドとなっているプロセスのみをKILLするので影響はない、らしい
            foreach (var p in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
            {
                if (p.MainWindowTitle == "")
                {
                    try { 
                        p.Kill();
                    }
                    // 既に終了 or 終了処理中 となっているプロセスをKILLしようとした場合、スルー
                    catch (InvalidOperationException e)
                    {

                    }
                }
            }
        }

        private void src_1_SelectedIndexChanged(object sender, EventArgs e)
        {
            set_combo_shibu(src_1.SelectedValue.ToString());
        }

        private String GetInsertKeikakuJouhou()
        {
            var SQLtext = new StringBuilder();

            SQLtext.AppendLine("INSERT INTO KeikakuJouhou(");
            SQLtext.AppendLine("KeikakuID");                            // 計画ID
            SQLtext.AppendLine(", KeikakuFileMei");                     // ファイル名
            SQLtext.AppendLine(", KeikakuSakuseibi");                   // 作成日
            SQLtext.AppendLine(", KeikakuSakuseiKashoShibuMei");        // 課所支部名
            SQLtext.AppendLine(", KeikakuTourokubi");                   // 計画登録日
            SQLtext.AppendLine(", KeikakuUriageNendo");                 // 売上年度
            SQLtext.AppendLine(", KeikakuBangou");                      // 計画番号
            SQLtext.AppendLine(", KeikakuZenkaiKeikakuBangou");         // 前年度計画番号
            SQLtext.AppendLine(", KeikakuZenkaiAnkenBangou");           // 前年度案件番号
            SQLtext.AppendLine(", KeikakuZenkaiJutakuBangou");          // 前年度受託番号
            SQLtext.AppendLine(", KeikakuZenkaiHachuushaMeiKaMei");     // 前年度発注者名・課名
            SQLtext.AppendLine(", KeikakuZenkaiGyoumuMei");             // 前年度業務名称
            SQLtext.AppendLine(", KeikakuBushoShibuCD");                // 部所支部CD
            SQLtext.AppendLine(", KeikakuBushoShibuMei");               // 部所支部名
            SQLtext.AppendLine(", KeikakuKashoShibuCD");                // 課所支部CD
            SQLtext.AppendLine(", KeikakuKashoShibuMei");               // 課所支部名
            SQLtext.AppendLine(", KeikakuGyoumuKubun");                 // 契約区分CD
            SQLtext.AppendLine(", KeikakuGyoumuKubunMei");              // 契約区分
            SQLtext.AppendLine(", KeikakuHachuushaMeiKaMei");           // 発注者名・課名
            SQLtext.AppendLine(", KeikakuAnkenMei");                    // 計画案件名
            SQLtext.AppendLine(", KeikakuRakusatsumikomigaku");         // 調査部配分　落札時見込額（税抜）
            SQLtext.AppendLine(", KeikakuHenkoumikomigaku");            // 調査部配分　変更見込額（税抜）
            SQLtext.AppendLine(", KeikakuMikomigaku");                  // 調査部配分　見込額合計（税抜）
            SQLtext.AppendLine(", KeikakuRakusatsumikomigakuJF");       // 事業普及部配分　落札時見込額（税抜）
            SQLtext.AppendLine(", KeikakuHenkoumikomigakuJF");          // 事業普及部配分　変更見込額（税抜）
            SQLtext.AppendLine(", KeikakuMikomigakuJF");                // 事業普及部配分　見込額合計（税抜）
            SQLtext.AppendLine(", KeikakuRakusatsumikomigakuJ");        // 情報システム部配分　落札時見込額（税抜）
            SQLtext.AppendLine(", KeikakuHenkoumikomigakuJ");           // 情報システム部配分　変更見込額（税抜）
            SQLtext.AppendLine(", KeikakuMikomigakuJ");                 // 情報システム部配分　見込額合計（税抜）
            SQLtext.AppendLine(", KeikakuRakusatsumikomigakuK");        // 総合研究所　落札時見込額（税抜）
            SQLtext.AppendLine(", KeikakuHenkoumikomigakuK");           // 総合研究所　変更見込額（税抜）
            SQLtext.AppendLine(", KeikakuMikomigakuK");                 // 総合研究所　見込額合計（税抜）
            SQLtext.AppendLine(", KeikakuMikomigakuGoukei");            // 見込額総合計(税抜）
            SQLtext.AppendLine(", KeikakuShizaiChousa");                // 調査業務別　配分　資材調査
            SQLtext.AppendLine(", KeikakuEizen");                       // 調査業務別　配分　営繕
            SQLtext.AppendLine(", KeikakuKikiruiChousa");               // 調査業務別　配分　機器類調査
            SQLtext.AppendLine(", KeikakuKoujiChousahi");               // 調査業務別　配分　工事費調査
            SQLtext.AppendLine(", KeikakuSanpaiChousa");                // 調査業務別　配分　産廃調査
            SQLtext.AppendLine(", KeikakuHokakeChousa");                // 調査業務別　配分　歩掛調査
            SQLtext.AppendLine(", KeikakuShokeihiChousa");              // 調査業務別　配分　諸経費調査
            SQLtext.AppendLine(", KeikakuGenkaBunseki");                // 調査業務別　配分　原価分析調査
            SQLtext.AppendLine(", KeikakuKijunsakusei");                // 調査業務別　配分　基準作成改訂
            SQLtext.AppendLine(", KeikakuKoukyouRoumuhi");              // 調査業務別　配分　公共労費調査
            SQLtext.AppendLine(", KeikakuRoumuhiKoukyouigai");          // 調査業務別　配分　労務費公共以外
            SQLtext.AppendLine(", KeikakuSonotaChousabu");              // 調査業務別　配分　その他調査部
            SQLtext.AppendLine(", KeikakuHaibunGoukei");                // 調査業務別　配分　合計
            SQLtext.AppendLine(", KeikakuAnkensu");                     // 案件数
            SQLtext.AppendLine(", KeikakuCreateDate");                  // 作成日時
            SQLtext.AppendLine(", KeikakuCreateUser");                  // 作成ユーザ
            SQLtext.AppendLine(", KeikakuCreateProgram");               // 作成機能
            SQLtext.AppendLine(", KeikakuUpdateDate");                  // 更新日時
            SQLtext.AppendLine(", KeikakuUpdateUser");                  // 更新ユーザ
            SQLtext.AppendLine(", KeikakuUpdateProgram");               // 更新機能
            SQLtext.AppendLine(", KeikakuDeleteFlag");                  // 削除フラグ
            SQLtext.AppendLine(")VALUES(");

            return SQLtext.ToString();
        }

        private String ChangeSQLString(object TextValue, int i, int j)
        {
            string SQLTextValue = "";

            if (TextValue != null)
            {
                SQLTextValue = GlobalMethod.ChangeSqlText(TextValue.ToString(), i, j);
            }

            return SQLTextValue;
        }

        private String ChangeSQLInt(object TextValue, int i, int j)
        {
            string SQLTextValue = "0";

            if (TextValue != null)
            {
                SQLTextValue = GlobalMethod.ChangeSqlText(TextValue.ToString(), i, j);
            }

            return SQLTextValue;
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
        // 更新履歴登録用SQLテキストの生成
        private string Insert_History_SQLText(string Naiyo, string ProgramNM)
        {
            string SQLText = "";
            String strSpace = " ";
            String strComma = ",";
            string strSingleQuote = "'";
            String strN = "N";
            String strCommaSpace = strComma + strSpace;
            StringBuilder sb = new StringBuilder();

            // キーの採番
            int Histry_No = GlobalMethod.getSaiban("HistoryID");

            // 値の加工
            string Today = strSingleQuote + DateTime.Today.ToString() + strSingleQuote;
            string KojinCD = strN + strSingleQuote + UserInfos[0] + strSingleQuote;
            string KojinNM = strN + strSingleQuote + UserInfos[1] + strSingleQuote;
            string BushoCD = strN + strSingleQuote + UserInfos[2] + strSingleQuote;
            string BushoNM = strN + strSingleQuote + UserInfos[3] + strSingleQuote;
            Naiyo = strN + strSingleQuote + Naiyo + strSingleQuote;
            ProgramNM = strSingleQuote + ProgramNM + strSingleQuote;

            // SQLテキストの作成
            sb.Clear();
            sb.Append("INSERT INTO T_HISTORY(");
            sb.Append("H_DATE_KEY");
            sb.Append(strCommaSpace);
            sb.Append("H_NO_KEY");
            sb.Append(strCommaSpace);
            sb.Append("H_OPERATE_DT");
            sb.Append(strCommaSpace);
            sb.Append("H_OPERATE_USER_ID");
            sb.Append(strCommaSpace);
            sb.Append("H_OPERATE_USER_MEI");
            sb.Append(strCommaSpace);
            sb.Append("H_OPERATE_USER_BUSHO_CD");
            sb.Append(strCommaSpace);
            sb.Append("H_OPERATE_USER_BUSHO_MEI");
            sb.Append(strCommaSpace);
            sb.Append("H_OPERATE_NAIYO");
            sb.Append(strCommaSpace);
            sb.Append("H_ProgramName");
            sb.Append(") VALUES (");
            sb.Append("SYSDATETIME()");     // H_DATE_KEY
            sb.Append(strCommaSpace);
            sb.Append(Histry_No);           // H_NO_KEY
            sb.Append(strCommaSpace);
            sb.Append(Today);               // H_OPERATE_DT
            sb.Append(strCommaSpace);
            sb.Append(KojinCD);             // H_OPERATE_USER_ID
            sb.Append(strCommaSpace);
            sb.Append(KojinNM);             // H_OPERATE_USER_MEI
            sb.Append(strCommaSpace);
            sb.Append(BushoCD);             // H_OPERATE_USER_BUSHO_CD
            sb.Append(strCommaSpace);
            sb.Append(BushoNM);             // H_OPERATE_USER_BUSHO_MEI
            sb.Append(strCommaSpace);
            sb.Append(Naiyo);               // H_OPERATE_NAIYO
            sb.Append(strCommaSpace);
            sb.Append(ProgramNM);           // H_ProgramName
            sb.Append(")");

            SQLText = sb.ToString();

            return SQLText;
        }

        private void btnGridSize_Click(object sender, EventArgs e)
        {
            //if (btnGridSize.Text == "一覧拡大")
            //{
            //    // height:691 → 1086・・・調査品目明細と合わせる
            //    // width:1864 → 3752
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid1.Height = 1086;
            //    c1FlexGrid1.Width = 3752;
            //}
            //else
            //{
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid1.Height = 691;
            //    c1FlexGrid1.Width = 1864;
            //}
            gridSizeChange();
        }
        private void gridSizeChange()
        {
            //if (btnGridSize.Text == "一覧拡大")
            //{
            //    // height:691 → 1086・・・調査品目明細と合わせる
            //    // width:1864 → 3752
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid1.Height = 1086;
            //    c1FlexGrid1.Width = 3752;
            //}
            //else
            //{
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid1.Height = 691;
            //    c1FlexGrid1.Width = 1864;
            //}
            string num = "";
            int bigHeight = 0;
            int bigWidth = 0;
            int smallHeight = 0;
            int smallWidth = 0;

            if (btnGridSize.Text == "一覧拡大")
            {
                num = GlobalMethod.GetCommonValue1("KEIKAKU_GRID_BIG_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out bigHeight);
                    if (bigHeight == 0)
                    {
                        bigHeight = 1086;
                    }
                }
                num = GlobalMethod.GetCommonValue1("KEIKAKU_GRID_BIG_WIDTH");
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
                num = GlobalMethod.GetCommonValue1("KEIKAKU_GRID_SMALL_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out smallHeight);
                    if (smallHeight == 0)
                    {
                        smallHeight = 691;
                    }
                }
                num = GlobalMethod.GetCommonValue1("KEIKAKU_GRID_SMALL_WIDTH");
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
