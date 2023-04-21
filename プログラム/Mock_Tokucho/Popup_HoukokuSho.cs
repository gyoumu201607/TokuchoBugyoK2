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

namespace TokuchoBugyoK2
{
    public partial class Popup_HoukokuSho : Form
    {
        public string[] UserInfos;
        public string[] ReturnValue = new string[10];
        GlobalMethod GlobalMethod = new GlobalMethod();
        public string MadoguchiID = "";
        public int MENU_ID = 203;
        public int PrintBunruiCD = 6;
        public string PrintGamen = "";
        public string BushoCD = "";
        public string Chousain = "";
        public string HoukokuSentaku = "0";
        public DateTime KikanStart;
        public DateTime KikanEnd;
        public string SeikyuuGetsu = "";
        public int ShuFuku = 0;
        public string Hinmei = "";
        public string Kikaku = "";
        public int Zaikou = 0;
        public int KuhakuList = 0;
        public DateTime Shimekiribi;
        public int HizukeKubun = 0;
        public string Memo1 = "";
        public string Memo2 = "";

        private Boolean changeCombo = false;
        private Boolean existFolder = false;
        private DateTime NullDate;

        Popup_Download download_form = null;

        public Popup_HoukokuSho()
        {
            InitializeComponent();

            // コンボボックスにマウスホイールイベントを付与
            this.comboBox_Nendo.MouseWheel += item_MouseWheel;
            this.comboBox_Month.MouseWheel += item_MouseWheel;
            this.comboBox_Quarter.MouseWheel += item_MouseWheel;
            this.comboBox_Chohyo.MouseWheel += item_MouseWheel;

        }

        private void Popup_HoukokuSho_Load(object sender, EventArgs e)
        {
            tableLayoutPanel3.Visible = false;
            // えんとり君修正STEP2 報告書共通化
            tableLayoutPanel5.Visible = false;

            // 単価契約画面からの遷移の場合
            if ("TankaKeiyaku".Equals(PrintGamen))
            {
                if (KikanStart != NullDate)
                {
                    dateTime_KikanStart.Value = KikanStart;
                }
                if (KikanEnd != NullDate)
                {
                    dateTime_KikanEnd.Value = KikanEnd;
                }
            }

            // コンボボックスの設定
            set_combo();

            // 値の取得
            get_data();

            // フォルダーパスのチェック
            FolderPathCheck();

            // 単価契約から呼び出した場合、DLを初期選択として変更不可にする
            if (MENU_ID == 208)
            {
                radioButton_DL.Checked = true;
                radioButton_Save.Enabled = false;
            }

            //  VIPS　20220330　課題管理表No1298(983)　ADD 自分大臣の時、初期選択はDLにする
            if (PrintGamen == "Jibun")
            {
                radioButton_DL.Checked = true;
            }
        }

        //コンボボックス設定
        private void set_combo()
        {
            // 各コンボのIndexChangedが動かないようにする。（ただし帳票のコンボはフラグを見ていないので動かす。）
            changeCombo = true;

            // 年度 コンボ
            string discript = "NendoSeireki";
            string value = "NendoID";
            string table = "Mst_Nendo";
            string where = "";
            //コンボボックスデータ取得
            DataTable combodt2 = GlobalMethod.getData(discript, value, table, where);
            comboBox_Nendo.DisplayMember = "Discript";
            comboBox_Nendo.ValueMember = "Value";
            comboBox_Nendo.DataSource = combodt2;
            comboBox_Nendo.SelectedValue = GlobalMethod.GetTodayNendo();

            // 月 コンボ
            DataTable combodt3 = new System.Data.DataTable();
            combodt3.Columns.Add("Value", typeof(int));
            combodt3.Columns.Add("Discript", typeof(string));
            combodt3.Rows.Add(0, "");
            combodt3.Rows.Add(4, "04");
            combodt3.Rows.Add(5, "05");
            combodt3.Rows.Add(6, "06");
            combodt3.Rows.Add(7, "07");
            combodt3.Rows.Add(8, "08");
            combodt3.Rows.Add(9, "09");
            combodt3.Rows.Add(10, "10");
            combodt3.Rows.Add(11, "11");
            combodt3.Rows.Add(12, "12");
            combodt3.Rows.Add(1, "01");
            combodt3.Rows.Add(2, "02");
            combodt3.Rows.Add(3, "03");
            comboBox_Month.DisplayMember = "Discript";
            comboBox_Month.ValueMember = "Value";
            comboBox_Month.DataSource = combodt3;

            // 四半期 コンボ
            // value の設定値は期の開始月とする。（後続の期間指定でそのまま利用するため）
            DataTable combodt4 = new System.Data.DataTable();
            combodt4.Columns.Add("Value", typeof(int));
            combodt4.Columns.Add("Discript", typeof(string));
            combodt4.Rows.Add(0, "");
            combodt4.Rows.Add(4, "第1期");
            combodt4.Rows.Add(7, "第2期");
            combodt4.Rows.Add(10, "第3期");
            combodt4.Rows.Add(1, "第4期");
            comboBox_Quarter.DisplayMember = "Discript";
            comboBox_Quarter.ValueMember = "Value";
            comboBox_Quarter.DataSource = combodt4;

            // 帳票選択 コンボ
            string BushoCD = "";
            discript = "ShuukeiMei";
            value = "ShuukeiMei";
            table = "Mst_Busho";
            where = "GyoumuBushoCD = '" + UserInfos[2] + "' AND ShuukeiMei = '1.本部' ";
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            if (combodt != null && combodt.Rows.Count > 0)
            {
                BushoCD = combodt.Rows[0][0].ToString();
            }
            else
            {
                BushoCD = UserInfos[2].ToString();
            }

            DataTable combodt1 = new System.Data.DataTable();
            discript = "PrintName";
            value = "PrintListID";
            table = "Mst_PrintList";
            // 20230131帳票出力性能改善対応：【共通】報告内容共通化
            //where = "MENU_ID = " + MENU_ID
            //        + " AND PrintBunruiCD = " + PrintBunruiCD
            //        + " AND (BushoKanriboBushoCD = '" + BushoCD + "' OR BushoKanriboBushoCD is null)"
            //        + " AND PrintDelFlg <> 1 ORDER BY PrintListNarabijun ";
            if (PrintGamen == "Tokumei")
            {
                where = "MENU_ID = " + MENU_ID
                      + " AND PrintBunruiCD = " + PrintBunruiCD
                      + " AND (BushoKanriboBushoCD = '" + BushoCD + "' OR BushoKanriboBushoCD is null)"
                      + " AND PrintHinagataBangou <> 21"
                      + " AND PrintDelFlg <> 1 ORDER BY PrintListNarabijun ";
            }
            else
            { 
                where = "MENU_ID = " + MENU_ID 
                      + " AND PrintBunruiCD = " + PrintBunruiCD
                      + " AND (BushoKanriboBushoCD = '" + BushoCD + "' OR BushoKanriboBushoCD is null)"
                      + " AND PrintDelFlg <> 1 ORDER BY PrintListNarabijun ";
            }
            Console.WriteLine(where);
            // コンボボックスデータ取得
            combodt1 = GlobalMethod.getData(discript, value, table, where);
            comboBox_Chohyo.DisplayMember = "Discript";
            comboBox_Chohyo.ValueMember = "Value";
            // ここだけはIndexChangedを動かすことで、各コンボの初期設定、表示切り替えを実施する。
            comboBox_Chohyo.DataSource = combodt1;

            changeCombo = false;

        }

        private void get_data()
        {
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            item1_HoukokuFolder.Text = "";
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();

                // 報告書フォルダ取得
                var dtCommon = new DataTable();
                cmd.CommandText = "SELECT MadoguchiHoukokuShoFolder"
                                + "  FROM MadoguchiJouhou "
                                + " WHERE MadoguchiID = '" + MadoguchiID + "'";
                // データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dtCommon);

                // 調査品目画面から呼び出された場合だけ、フォルダを設定する。（単価契約画面から呼ばれた場合、セットしない）
                //if (dtCommon.Rows.Count > 0)
                //if (dtCommon.Rows.Count > 0 && MENU_ID == 203)
                if (dtCommon.Rows.Count > 0 && MENU_ID != 208)
                {
                    item1_HoukokuFolder.Text = dtCommon.Rows[0][0].ToString();
                }

                conn.Close();
            }
            set_data();
            //Resize_Grid("c1FlexGrid1");
        }

        private void set_data()
        {
            // 特に処理なし
        }

        // プリントマスタの期間指定フラグの取得
        private Boolean get_PrintKikanFlg()
        {
            string PrintKikanFlg = get_Mst_Print("PrintKikanFlg");

            if (string.IsNullOrEmpty(PrintKikanFlg) || PrintKikanFlg == "0")
            {
                return false;
            }
            else
            {
                return true;
            }

        }

        // プリントマスタの雛型番号の取得
        private string get_PrintHinagataBangou()
        {
            string PrintHinagataBangou = get_Mst_Print("PrintHinagataBangou");

            return PrintHinagataBangou;
        }

        // プリントマスタのファイル名の取得
        private string get_PritFileName()
        {
            string PrintFileName = get_Mst_Print("PrintDownloadFileName");

            return PrintFileName;
        }

        // プリントマスタから指定した項目の値を取得する
        private string get_Mst_Print(string targetItem)
        {
            string returnValue = "";

            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            using (var conn = new SqlConnection(connStr))
            {
                if (comboBox_Chohyo.SelectedValue != null)
                {
                    var cmd = conn.CreateCommand();

                    // プリントマスタの値の取得
                    var dtprint = new DataTable();
                    cmd.CommandText = "SELECT " + targetItem
                                    + "  FROM Mst_PrintList "
                                    + " WHERE PrintListID = " + comboBox_Chohyo.SelectedValue + "";
                    // データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dtprint);
                    if (dtprint.Rows.Count > 0)
                    {
                        returnValue = dtprint.Rows[0][0].ToString();
                    }
                }
                conn.Close();
            }
            return returnValue;

        }

        private void button_end_Click(object sender, EventArgs e)
        {
            // 調査品目データを取り直しさせるためにパラメータをセット
            ReturnValue[0] = "1";
            this.Close();
        }

        private void folderHoukokushoIcon_Click(object sender, EventArgs e)
        {
            if (item1_HoukokuFolder.Text != "")
            {
                // ファイルパスとして認識できる場合のみ、エクスプローラーで表示する
                if (System.Text.RegularExpressions.Regex.IsMatch(item1_HoukokuFolder.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    // 指定されたフォルダパスが存在するなら開く
                    if (item1_HoukokuFolder.Text != "" && item1_HoukokuFolder.Text != null && Directory.Exists(item1_HoukokuFolder.Text))
                    {
                        System.Diagnostics.Process.Start(GlobalMethod.GetPathValid(item1_HoukokuFolder.Text));
                        return;
                    }
                }
            }
            System.Diagnostics.Process.Start("EXPLORER.EXE", "");
        }

        private void FolderPathCheck()
        {
            // 報告書フォルダ
            if (Directory.Exists(item1_HoukokuFolder.Text))
            {
                item1_Folder_icon.Image = Image.FromFile("Resource/Image/folder_yellow_s.png");

                // ファイル出力ボタンを活性化
                btnFileExport.Enabled = true;

                set_error("", 0);
                existFolder = true;
            }
            else
            {
                item1_Folder_icon.Image = Image.FromFile("Resource/Image/folder_gray_s.png");

                // ファイル出力ボタンを非活性化
                btnFileExport.Enabled = false;
                btnFileExport.BackColor = Color.DimGray;

                set_error("", 0);
                set_error(GlobalMethod.GetMessage("E20333", ""));
                existFolder = false;
            }
        }

        // ファイル出力
        private void btnFileExport_Click(object sender, EventArgs e)
        {
            // ボタン押下時にボタンを非活性にして押せなくする。
            btnFileExport.BackColor = Color.DimGray;
            btnFileExport.Enabled = false;

            // 0:MadoguchiID     窓口ID
            // 1:PrintGamen      呼び出し元画面 0:窓口ミハル 1:特命課長 2:自分大臣 3:単価契約
            // 2:DateFrom        日付From
            // 3:DateTo          日付To
            // 4:BushoCD         部所CD
            // 5:Tantousha       担当者名
            // 6:HoukokuSentaku  日付選択
            // 7:seikyuuGetsu    請求月
            // 8:ShuFuku         主副 0:主＋副 1:主 2:副
            // 9:Hinmei          品名
            // 10:Kikaku         規格
            // 11:Zaikou         材工 0:全て 1:材のみ 2:工のみ 3:材+D工 4:E工のみ 5:他
            // 12:KuhakuList     担当者空白リスト 0:全て 1:担当者が空白のリスト 2:担当者が設定済のリスト
            // 13:Shimekiribi    締切日
            // 14:HizukeKubun    日付区分 0:空 1:前日 2:当日 3:翌日
            // 15:Memo1          メモ１
            // 16:Memo2          メモ２
            // 17:ChuushiYouhi   品目の中止を含む ※報告書共通化出力のみ

            // 17個分先に用意
            string[] report_data = new string[17] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };


            // 窓口ID
            report_data[0] = MadoguchiID;

            // 呼び出し元画面
            report_data[1] = "0";
            switch (PrintGamen)
            {
                case "Madoguchi":
                    report_data[1] = "0";
                    break;
                case "Tokumei":
                    report_data[1] = "1";
                    break;
                case "Jibun":
                    report_data[1] = "2";
                    break;
                default:
                    break;
            }
            // 日付From
            report_data[2] = "null";
            if (dateTime_KikanStart.CustomFormat == "")
            {
                report_data[2] = "'" + dateTime_KikanStart.Text + "'";
            }
            // 日付To
            report_data[3] = "null";
            if (dateTime_KikanEnd.CustomFormat == "")
            {
                report_data[3] = "'" + dateTime_KikanEnd.Text + "'";
            }
            // 部所CD
            report_data[4] = BushoCD;
            // 担当者名
            report_data[5] = Chousain;
            // 日付選択
            report_data[6] = HoukokuSentaku;
            // 請求月
            report_data[7] = SeikyuuGetsu;
            // 主副
            report_data[8] = ShuFuku.ToString();
            // 品名
            report_data[9] = Hinmei;
            // 規格
            report_data[10] = Kikaku;
            // 材工
            report_data[11] = Zaikou.ToString();
            // 担当者空白リスト
            report_data[12] = KuhakuList.ToString();
            // 締切日
            report_data[13] = "null";
            if (Shimekiribi != NullDate)
            {
                report_data[13] = "'" + Shimekiribi.ToString() + "'";
            }
            // 日付区分
            report_data[14] = HizukeKubun.ToString();
            // メモ１
            report_data[15] = Memo1;
            // メモ２
            report_data[16] = Memo2;

            int listID = 0;
            if (comboBox_Chohyo.SelectedValue == null)
            {
                // 雛型ファイルが存在しません。
                set_error("", 0);
                set_error(GlobalMethod.GetMessage("E20112", ""));
                return;
            }
            else
            {
                listID = int.Parse(comboBox_Chohyo.SelectedValue.ToString());
            }

            // No.1423 1197 報告書共通化の報告書追加時に条件が表示されない。
            string printDataPattern = "";
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                try
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();
                    var Dt = new System.Data.DataTable();
                    //SQL生成
                    cmd.CommandText = "SELECT " +
                      "PrintDataPattern,PrintKikanFlg " +
                      "FROM " + "Mst_PrintList " +
                      "WHERE PrintListID = '" + listID + "'";

                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(Dt);
                    //Boolean errorFLG = false;

                    if (Dt.Rows.Count > 0 && (Dt.Rows[0][0].ToString() == "800" || Dt.Rows[0][0].ToString() == "801"))
                    {
                        printDataPattern = Dt.Rows[0][0].ToString();
                        Array.Resize(ref report_data, 18);
                        report_data[17] = "null";
                        if (radioButton_No.Checked)
                        {
                            report_data[17] = "0";
                        }
                        else
                        {
                            report_data[17] = "1";
                        }
                    }
                }
                catch (Exception)
                {
                    // 何もしない
                }
                finally
                {
                    conn.Close();
                }

            }
            // えんとり君修正STEP2 報告書共通化
            //if (listID == 802 || listID == 803)
            //{
            //    Array.Resize(ref report_data, 18);
            //    report_data[17] = "null";
            //    if (radioButton_No.Checked)
            //    {
            //        report_data[17] = "0";
            //    }
            //    else
            //    {
            //        report_data[17] = "1";
            //    }
            //}
            //// ファイル名の取得
            //string w_PritFileName = get_PritFileName();

            string printName = "";
            if (listID == 233)
            {
                printName = "TantoushaIchiran";
            }
            else
            {
                printName = "Houkokusho";
            }

            string[] result = GlobalMethod.InsertMadoguchiReportWork(listID, UserInfos[0], report_data, printName, printDataPattern);
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
                    if (result[1] == "")
                    {
                        set_error(GlobalMethod.GetMessage("E00091", ""));
                    }
                    else
                    {
                        set_error(result[1]);
                    }
                }
                else
                {
                    set_error("", 0);

                    // 直接フォルダに保存するかDLダイアログを表示するか選択させる
                    if (radioButton_Save.Checked)
                    {
                        // 成功時は、ファイルをフォルダにコピーする
                        try
                        {
                            //System.IO.File.Copy(result[2], item1_HoukokuFolder.Text + @"\" + w_PritFileName, true);
                            //set_error("報告書ファイルを出力しました。:" + w_PritFileName);
                            System.IO.File.Copy(result[2], item1_HoukokuFolder.Text + @"\" + result[3], true);
                            set_error("報告書ファイルを出力しました。:" + result[3]);

                            // 調査品目データを取り直しさせるためにパラメータをセット
                            ReturnValue[0] = "1";
                        }
                        catch (Exception)
                        {
                            // ファイルコピー失敗
                            set_error(GlobalMethod.GetMessage("E20334", ""));
                        }
                    }
                    else
                    {
                        if (download_form != null)
                        {
                            download_form.Close();
                        }
                        // DLダイアログを表示する。
                        download_form = new Popup_Download();
                        download_form.TopLevel = false;
                        this.Controls.Add(download_form);

                        String fileName = Path.GetFileName(result[3]);
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

            // 処理終了後にボタンを使用できるように活性化する
            btnFileExport.Enabled = true;
            btnFileExport.BackColor = Color.FromArgb(42, 78, 122);

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

        // コンボボックス変更時（帳票）
        private void comboBox_ChohyoSelectedIndexChanged(object sender, EventArgs e)
        {
            // 期間指定の表示の切り替え
            string w_PrintHinagataBangou = get_PrintHinagataBangou();
            if (get_PrintKikanFlg() || w_PrintHinagataBangou == "28" || w_PrintHinagataBangou == "22")
            {
                tableLayoutPanel3.Visible = true;

                changeCombo = false;

                // 初期値の設定
                switch (w_PrintHinagataBangou)
                {
                    case "22":  // 報告書・事業所
                        comboBox_Month.SelectedValue = DateTime.Today.Month.ToString();     // 月    　
                        comboBox_Quarter.SelectedValue = 0;                                 // 四半期
                        setkikan(int.Parse(comboBox_Month.SelectedValue.ToString()), 1);    // 期間
                        break;
                    case "28":  // 報告書・農政
                        comboBox_Month.SelectedValue = 0;                                   // 月
                        comboBox_Quarter.SelectedValue = 4;                                 // 四半期
                        setkikan(int.Parse(comboBox_Quarter.SelectedValue.ToString()), 3);  // 期間
                        break;
                    case "36":  // 報告書・北陸地整-調査区分内訳書
                        comboBox_Month.SelectedValue = DateTime.Today.Month.ToString();     // 月
                        comboBox_Quarter.SelectedValue = 0;                                 // 四半期
                        setkikan(int.Parse(comboBox_Month.SelectedValue.ToString()), 1);    // 期間
                        break;
                    case "37":  // 報告書・北陸地整-報告用調査区分確認簿
                        comboBox_Month.SelectedValue = 0;                                   // 月
                        comboBox_Quarter.SelectedValue = 4;                                 // 四半期
                        setkikan(int.Parse(comboBox_Quarter.SelectedValue.ToString()), 3);  // 期間
                        break;
                    case "38":  // 報告書・北陸地整-隣接地域調査報告書
                        comboBox_Month.SelectedValue = DateTime.Today.Month.ToString();     // 月
                        comboBox_Quarter.SelectedValue = 0;                                 // 四半期
                        setkikan(int.Parse(comboBox_Month.SelectedValue.ToString()), 1);    // 期間
                        break;
                    case "39":  // 報告書・北陸地整-契約内訳書ランク別
                        comboBox_Month.SelectedValue = 0;                                   // 月
                        comboBox_Quarter.SelectedValue = 4;                                 // 四半期
                        setkikan(int.Parse(comboBox_Quarter.SelectedValue.ToString()), 3);  // 期間
                        break;
                    default:
                        comboBox_Month.SelectedValue = DateTime.Today.Month.ToString();     // 月
                        comboBox_Quarter.SelectedValue = 0;                                 // 四半期
                        setkikan(int.Parse(comboBox_Month.SelectedValue.ToString()), 1);    // 期間
                        break;
                }

            }
            else
            {
                tableLayoutPanel3.Visible = false;
            }

            // えんとり君修正STEP2 報告書共通化
            if(w_PrintHinagataBangou == "21")
            {
                tableLayoutPanel5.Visible = true;
            }
            else
            {
                tableLayoutPanel5.Visible = false;
            }
        }

        // コンボボックス変更時（年度）
        private void comboBox_NendoSelectedIndexChanged(object sender, EventArgs e)
        {
            changeMonthQuarter("Nendo");
        }

        // コンボボックス変更時（月）
        private void comboBox_MonthSelectedIndexChanged(object sender, EventArgs e)
        {
            changeMonthQuarter("Month");
        }

        // コンボボックス変更時（四半期）
        private void comboBox_QuarterSelectedIndexChanged(object sender, EventArgs e)
        {
            changeMonthQuarter("Quarter");
        }

        private void changeMonthQuarter(string combo)
        {

            // コンボの値を変更することでIndexChangedが動いてしまうため、フラグで制御を行う。
            if (!changeCombo)
            {
                changeCombo = true;

                // 年度のコンボの変更後の値によって、各コンボの値を変更する。
                if (combo == "Nendo")
                {
                    // 月と四半期のコンボの値によって、期間を設定する。
                    if (comboBox_Month.SelectedValue.ToString() != "0")
                    {
                        setkikan(int.Parse(comboBox_Month.SelectedValue.ToString()), 1);        // 期間
                    }
                    else
                    {
                        if (comboBox_Quarter.SelectedValue.ToString() != "0")
                        {
                            setkikan(int.Parse(comboBox_Quarter.SelectedValue.ToString()), 3);  // 期間
                        }
                    }
                }

                // 月のコンボの変更後の値によって、各コンボの値を変更する。
                if (combo == "Month")
                {
                    if (comboBox_Month.SelectedValue.ToString() == "0")
                    {
                        comboBox_Month.SelectedValue = DateTime.Today.Month.ToString();     // 月
                        comboBox_Quarter.SelectedValue = 0;                                 // 四半期
                    }
                    else
                    {
                        comboBox_Quarter.SelectedValue = 0;                                 // 四半期
                    }
                    setkikan(int.Parse(comboBox_Month.SelectedValue.ToString()), 1);        // 期間
                }

                // 四半期のコンボの変更後の値によって、各コンボの値を変更する。
                if (combo == "Quarter")
                {
                    if (comboBox_Quarter.SelectedValue.ToString() == "0")
                    {
                        comboBox_Month.SelectedValue = DateTime.Today.Month.ToString();     // 月
                        setkikan(int.Parse(comboBox_Month.SelectedValue.ToString()), 1);    // 期間
                    }
                    else
                    {
                        comboBox_Month.SelectedValue = 0;                                   // 月
                        setkikan(int.Parse(comboBox_Quarter.SelectedValue.ToString()),3);   // 期間
                    }
                }

                changeCombo = false;
            }

        }

        // 期間の設定（受け取った月（month）を起点に指定した期間（kikan）でFrom,Toを設定する）
        private void setkikan(int month, int kikan)
        {
            // 期間の設定
            int.TryParse(comboBox_Nendo.SelectedValue.ToString(), out int year);
            DateTime firstDay = new DateTime(year, month, 1);
            if (month <= 3)
            {
                firstDay = firstDay.AddYears(1);
            }
            DateTime lastDay = firstDay.AddMonths(kikan).AddDays(-1);

            // 期間From
            dateTime_KikanStart.CustomFormat = "";
            dateTime_KikanStart.Text = firstDay.ToString();

            // 期間To
            dateTime_KikanEnd.CustomFormat = "";
            dateTime_KikanEnd.Text = lastDay.ToString();

        }

        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        private void radioButton_DL_CheckedChanged(object sender, EventArgs e)
        {
            // DLが選択されている時はフォルダ有無に関係なく出力できるようにする
            if (radioButton_DL.Checked || existFolder)
            {
                // ファイル出力ボタンを活性化
                btnFileExport.Enabled = true;
                btnFileExport.BackColor = Color.FromArgb(42, 78, 122);
            }
            else
            {
                // ファイル出力ボタンを非活性化
                btnFileExport.Enabled = false;
                btnFileExport.BackColor = Color.DimGray;
            }
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

    }
}
