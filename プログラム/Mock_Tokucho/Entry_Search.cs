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
using C1.Win.C1Input;
using System.Configuration;
using System.Collections;
using Microsoft.VisualBasic.ApplicationServices;
using System.IO;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace TokuchoBugyoK2
{
    public partial class Entry_Search : Form
    {
        public string[] UserInfos;
        private DataTable ListData = new DataTable();
        public Boolean ReSearch = true;
        GlobalMethod GlobalMethod = new GlobalMethod();
        public Entry_Search()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            //ユーザ名を設定
            label3.Text = UserInfos[3] + "：" + UserInfos[1];

            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            //一覧のヘッダーをマージするために、2行目にコピー
            for (int i = 0; i < c1FlexGrid1.Cols.Count; i++)
            {
                c1FlexGrid1[1, i] = c1FlexGrid1[0, i];
            }

            

            gridSizeChange();

            // ホイール制御
            this.item1_KoukiNendo.MouseWheel += item_MouseWheel; // 工期開始年度
            this.src_4.MouseWheel += item_MouseWheel; // 受託課所支部
            this.src_5.MouseWheel += item_MouseWheel; // 事業部
            this.src_6.MouseWheel += item_MouseWheel; // 案件区分
            this.src_1.MouseWheel += item_MouseWheel; // 売上年度
            this.src_10.MouseWheel += item_MouseWheel; // 入札状況
            this.src_12.MouseWheel += item_MouseWheel; // 発注者区分1
            this.src_9.MouseWheel += item_MouseWheel; // 契約区分
            this.src_21.MouseWheel += item_MouseWheel; // 受注意欲
            this.src_22.MouseWheel += item_MouseWheel; // 引合状況
            this.src_23.MouseWheel += item_MouseWheel; // 当会応札
            this.src_20.MouseWheel += item_MouseWheel; // 参考見積
            this.src_26.MouseWheel += item_MouseWheel; // 起案状況
            this.src_28.MouseWheel += item_MouseWheel; // 表示件数
            this.comboBox13.MouseWheel += item_MouseWheel; // 帳票 

            // コントロールを初期化します。
            c1FlexGrid1.Styles.Normal.WordWrap = true;
            c1FlexGrid1.Rows[0].AllowMerging = true;
            c1FlexGrid1.AllowAddNew = false;

            // 部門ごとのヘッダーをマージ
            //C1.Win.C1FlexGrid.CellRange rng = c1FlexGrid1.GetCellRange(0, 31, 0, 33);
            //rng.Data = "調査部部門";
            //rng = c1FlexGrid1.GetCellRange(0, 37, 0, 39);
            //rng.Data = "事業普及部部門";
            //rng = c1FlexGrid1.GetCellRange(0, 40, 0, 42);
            //rng.Data = "情報システム部部門";
            //rng = c1FlexGrid1.GetCellRange(0, 43, 0, 45);
            //rng.Data = "総合研究所部門";
            C1.Win.C1FlexGrid.CellRange rng = c1FlexGrid1.GetCellRange(0, 32, 0, 34);
            rng.Data = "調査部部門";
            rng = c1FlexGrid1.GetCellRange(0, 38, 0, 40);
            rng.Data = "事業普及部部門";
            rng = c1FlexGrid1.GetCellRange(0, 41, 0, 43);
            rng.Data = "情報システム部部門";
            rng = c1FlexGrid1.GetCellRange(0, 44, 0, 46);
            rng.Data = "総合研究所部門";

            //コンボボックスの内容を設定
            set_combo();
            ClearForm();

            //一覧に表示するデータを取得
            get_date();

            //ソート項目にアイコンを設定
            C1.Win.C1FlexGrid.CellRange cr;
            Bitmap bmp1 = new Bitmap("Resource/Image/SortIconDefalt.png");
            Bitmap bmpSort = new Bitmap(bmp1, bmp1.Width / 6, bmp1.Height / 6);
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
            cr = c1FlexGrid1.GetCellRange(0, 11);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 12);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 14);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 17);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 18);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 19);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 20);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 21);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 25);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 26);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 27);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 28);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 29);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 47);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            cr = c1FlexGrid1.GetCellRange(0, 50);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 6);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 7);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 8);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 10);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 11);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 13);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 16);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 17);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 18);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 19);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 20);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 24);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 26);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 27);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            ////cr = c1FlexGrid1.GetCellRange(0, 28);
            ////cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            ////cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 46);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid1.GetCellRange(0, 49);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
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

        // Gridの編集セルをクリックした場合の動作
        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //GlobalMethod.outputLogger("entory表示", "表示開始 " + DateTime.Now, "GetAnkenJouhou", UserInfos[1]);

            var hti = this.c1FlexGrid1.HitTest(new System.Drawing.Point(e.X, e.Y));
            var _row = hti.Row;
            var _col = hti.Column;

            if (hti.Column == 2 & hti.Row > 1 && c1FlexGrid1[hti.Row, hti.Column].ToString() == "1")
            {
                string mode;
                mode = "view";

                // Role:2システム管理者
                if (UserInfos[4].Equals("2"))
                {
                    mode = "";
                }
                if (UserInfos[2].Equals(c1FlexGrid1[hti.Row, 21].ToString()))
                {
                    mode = "";
                }
                else
                {
                    if (UserInfos[2].Substring(0, 4).Equals("1284") && (c1FlexGrid1[hti.Row, 21].ToString() == "127800" || c1FlexGrid1[hti.Row, 21].ToString() == "127910"))
                    {
                        mode = "";
                    }
                    if (UserInfos[2].Substring(0, 4).Equals("1292"))
                    {
                        mode = "";
                    }
                    if (UserInfos[2].Substring(0, 4).Equals("1502"))
                    {
                        mode = "";
                    }
                    if (UserInfos[2].Substring(0, 4).Equals("1504"))
                    {
                        mode = "";
                    }
                }
                this.ReSearch = true;
                string AnkenID = c1FlexGrid1[hti.Row, hti.Column + 1].ToString();
                //GlobalMethod.outputLogger("entory new", "new 開始 " + DateTime.Now, "GetAnkenJouhou", UserInfos[1]);
                Entry_Input form = new Entry_Input();
                //GlobalMethod.outputLogger("entory new", "new 終了 " + DateTime.Now, "GetAnkenJouhou", UserInfos[1]);
                form.mode = mode;
                form.AnkenID = AnkenID;
                form.UserInfos = this.UserInfos;
                //GlobalMethod.outputLogger("form.Show", "form.Show 開始 " + DateTime.Now, "GetAnkenJouhou", UserInfos[1]);
                form.Show(this);
                //GlobalMethod.outputLogger("form.Show", "form.Show 終了 " + DateTime.Now, "GetAnkenJouhou", UserInfos[1]);
            }
            // 契約図書
            if (hti.Column == 15 & hti.Row > 1)
            {
                if (c1FlexGrid1[hti.Row, 15].ToString() == "1")
                {
                    System.Diagnostics.Process.Start("EXPLORER.EXE", GlobalMethod.GetPathValid(c1FlexGrid1[hti.Row, 16].ToString()));
                }
            }

            //GlobalMethod.outputLogger("entory表示", "表示終了 " + DateTime.Now, "GetAnkenJouhou", UserInfos[1]);
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        //「新規」ボタン押下処理
        private void button2_Click(object sender, EventArgs e)
        {
            this.ReSearch = true;
            Entry_Input form = new Entry_Input();
            form.mode = "insert";
            form.UserInfos = this.UserInfos;

            form.Show(this);
        }

        //「変更伝票」ボタン押下処理
        private void button3_Click(object sender, EventArgs e)
        {
            Boolean checkflg = true;
            string checkNo = "";
            string busho = "";
            string saishin = "";
            string ankenkbn = "";
            string kianflg = "";
            set_error("", 0);
            for (int i = 2; i < c1FlexGrid1.Rows.Count; i++)
            {
                if (c1FlexGrid1.GetCellCheck(i, 1) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                {
                    if (checkNo == "")
                    {
                        checkNo = c1FlexGrid1[i, 3].ToString();
                        busho = c1FlexGrid1[i, 21].ToString();
                        saishin = c1FlexGrid1[i, 53].ToString();
                        ankenkbn = c1FlexGrid1[i, 19].ToString();
                        kianflg = c1FlexGrid1[i, 17].ToString();
                    }
                }
            }
            if (checkNo == "")
            {
                // E10002:契約情報を選択してください。
                set_error(GlobalMethod.GetMessage("E10002", ""));
                checkflg = false;
            }
            else
            {
                // 1:管理職ロール で、選択した部所が異なる場合は権限エラー
                if ("1".Equals(UserInfos[4]) && !busho.Equals(UserInfos[2]))
                {
                    // E10003:閲覧権限しかない為、変更伝票は作成できません。
                    set_error(GlobalMethod.GetMessage("E10003", ""));
                    checkflg = false;
                }
                if (saishin != "1")
                {
                    // E10006:最新伝票を選択してください。
                    set_error(GlobalMethod.GetMessage("E10006", ""));
                    checkflg = false;
                }
                if (ankenkbn.Equals("02") || ankenkbn.Equals("04") || ankenkbn.Equals(""))
                {
                    // E10007:契約変更(赤伝)、中止に対し、変更伝票は作成できません。
                    set_error(GlobalMethod.GetMessage("E10007", ""));
                    checkflg = false;
                }
                if (kianflg == "False")
                {
                    // E10008:起案されていないので、作成できません。
                    set_error(GlobalMethod.GetMessage("E10008", ""));
                    checkflg = false;
                }
            }

            if (checkflg)
            {
                this.ReSearch = true;
                Entry_Input form = new Entry_Input();
                form.mode = "change";
                form.AnkenID = checkNo;
                form.UserInfos = UserInfos;
                form.Show(this);
            }
        }


        //ヘッダー「計画」ボタン押下処理
        private void button6_Click(object sender, EventArgs e)
        {
            Entry_keikaku_Search form = new Entry_keikaku_Search();
            form.UserInfos = UserInfos;
            form.Show();
            this.Close();
        }

        //Grid「管理月」変更後の処理
        private void c1FlexGrid1_AfterEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            if (e.Col == 11)
            {
                set_error("", 0);
                string edit = c1FlexGrid1.Cols[e.Col][e.Row].ToString();
                if (edit != "")
                {
                    if (System.Text.RegularExpressions.Regex.IsMatch(edit, @"^\d{4}/\d{2}$"))
                    {
                        Boolean errorflg = true;
                        int year;
                        int month;
                        if (int.TryParse(edit.Substring(0, 4), out year) && int.TryParse(edit.Substring(5, 2), out month))
                        {
                            if ((year < 1900 || year > 2100) || (month < 1 || month > 12))
                            {
                                set_error(GlobalMethod.GetMessage("E10004", ""));
                                errorflg = false;
                            }
                        }
                        else
                        {
                            set_error(GlobalMethod.GetMessage("E10004", ""));
                            errorflg = false;
                        }
                        if (errorflg)
                        {
                            string naiyo = "ID:" + c1FlexGrid1.Rows[e.Row][3].ToString() + " 受託番号:" + c1FlexGrid1.Rows[e.Row][8].ToString() + " 管理月:" + edit;
                            if (GlobalMethod.Check_Table(c1FlexGrid1.Rows[e.Row][3].ToString(), "AnkenJouhouID", "AnkenJouhou", ""))
                            {
                                GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報 管理月更新 " + naiyo, "Update_Entry_Kanrizuki", "");
                                string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                                SqlConnection sqlconn = new SqlConnection(connStr);
                                sqlconn.Open();
                                SqlTransaction transaction = sqlconn.BeginTransaction();
                                var cmd = sqlconn.CreateCommand();
                                cmd.Transaction = transaction;
                                try
                                {
                                    cmd.CommandText = "UPDATE AnkenJouhou SET  " +
                                        "AnkenUriagebi = '" + edit + "/01" + "' " +
                                        ",AnkenUpdateDate = SYSDATETIME() " +
                                        ",AnkenUpdateUser = N'" + UserInfos[0] + "' " +
                                        ",AnkenUpdateProgram = 'Update_Entry_Kanrizuki' " +
                                         " WHERE AnkenJouhouID = '" + c1FlexGrid1.Rows[e.Row][3].ToString() + "' ";
                                    Console.WriteLine(cmd.CommandText);
                                    cmd.ExecuteNonQuery();
                                    transaction.Commit();

                                }
                                catch
                                {
                                    transaction.Rollback();
                                    throw;
                                }
                                finally
                                {
                                    sqlconn.Close();
                                }
                            }
                            else
                            {
                                GlobalMethod.outputLogger("Update_Entry_Kanrizuki", "契約情報 管理月更新対象なし " + naiyo, "GetAnkenJouhou", UserInfos[1]);
                            }
                            C1.Win.C1FlexGrid.CellRange rg;
                            rg = c1FlexGrid1.GetCellRange(e.Row, e.Col);
                            rg.Style = c1FlexGrid1.Styles["Editor"];
                        }
                        else
                        {
                            C1.Win.C1FlexGrid.CellRange rg;
                            rg = c1FlexGrid1.GetCellRange(e.Row, e.Col);
                            c1FlexGrid1.Styles["ErrorStyle"].BackColor = Color.FromArgb(255, 204, 255);
                            rg.Style = c1FlexGrid1.Styles["ErrorStyle"];
                        }
                    }
                    else
                    {
                        set_error(@"管理月はYYYY/MM形式で入力してください。");
                        C1.Win.C1FlexGrid.CellRange rg;
                        rg = c1FlexGrid1.GetCellRange(e.Row, e.Col);
                        c1FlexGrid1.Styles["ErrorStyle"].BackColor = Color.FromArgb(255, 204, 255);
                        rg.Style = c1FlexGrid1.Styles["ErrorStyle"];
                    }
                }
                // 管理月入力×が空だった場合
                else
                {
                    string naiyo = "ID:" + c1FlexGrid1.Rows[e.Row][3].ToString() + " 受託番号:" + c1FlexGrid1.Rows[e.Row][8].ToString() + " 管理月:" + edit;
                    if (GlobalMethod.Check_Table(c1FlexGrid1.Rows[e.Row][3].ToString(), "AnkenJouhouID", "AnkenJouhou", ""))
                    {
                        GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報 管理月更新 " + naiyo, "Update_Entry_Kanrizuki", "");
                        string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                        SqlConnection sqlconn = new SqlConnection(connStr);
                        sqlconn.Open();
                        SqlTransaction transaction = sqlconn.BeginTransaction();
                        var cmd = sqlconn.CreateCommand();
                        cmd.Transaction = transaction;
                        try
                        {
                            cmd.CommandText = "UPDATE AnkenJouhou SET  " +
                                "AnkenUriagebi = null " + // 管理月入力を空にする
                                ",AnkenUpdateDate = SYSDATETIME() " +
                                ",AnkenUpdateUser = N'" + UserInfos[0] + "' " +
                                ",AnkenUpdateProgram = 'Update_Entry_Kanrizuki' " +
                                 " WHERE AnkenJouhouID = '" + c1FlexGrid1.Rows[e.Row][3].ToString() + "' ";
                            Console.WriteLine(cmd.CommandText);
                            cmd.ExecuteNonQuery();
                            transaction.Commit();

                        }
                        catch
                        {
                            transaction.Rollback();
                            throw;
                        }
                        finally
                        {
                            sqlconn.Close();
                        }
                    }

                    C1.Win.C1FlexGrid.CellRange rg;
                    rg = c1FlexGrid1.GetCellRange(e.Row, e.Col);
                    rg.Style = c1FlexGrid1.Styles["Editor"];
                }
            }

        }


        private void dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            ((DateTimePicker)sender).CustomFormat = "";
        }

        private void dateTimePicker_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Delete)
            {
                ((DateTimePicker)sender).CustomFormat = " ";
            }
        }


        private void set_combo()
        {

            //コンボボックスの内容を設定
            GlobalMethod GlobalMethod = new GlobalMethod();
            //売上年度
            String discript = "NendoSeireki";
            String value = "NendoID ";
            String table = "Mst_Nendo";

            // 売上年度は年度マスタのデータを全て表示
            //String where = "NendoID <= YEAR(GETDATE()) AND NendoID > YEAR(GETDATE()) - 3 ORDER BY NendoSeireki DESC";
            String where = "";
            //コンボボックスデータ取得
            DataTable nendoCombodt = GlobalMethod.getData(discript, value, table, where);
            // VIPS 20220221 課題管理表No.1279(973) DEL プルダウンの空白が不要
            //DataRow nendodr;
            //if (nendoCombodt != null)
            //{
            //    nendodr = nendoCombodt.NewRow();
            //    nendoCombodt.Rows.InsertAt(nendodr, 0);
            //}
            src_1.DataSource = nendoCombodt;
            src_1.DisplayMember = "Discript";
            src_1.ValueMember = "Value";

            DataTable koukiCombodt = GlobalMethod.getData(discript, value, table, where);
            // VIPS 20220221 課題管理表No.1279(973) DEL プルダウンの空白が不要
            //DataRow koukinendodr;
            //// 空行追加
            //if (koukiCombodt != null)
            //{
            //    koukinendodr = koukiCombodt.NewRow();
            //    koukiCombodt.Rows.InsertAt(koukinendodr, 0);
            //}
            item1_KoukiNendo.DataSource = koukiCombodt;
            item1_KoukiNendo.DisplayMember = "Discript";
            item1_KoukiNendo.ValueMember = "Value";

            //受託課所支部
            discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
            value = "Mst_Busho.GyoumuBushoCD ";
            table = "Mst_Busho";
            where = "";
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            SortedList sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[20].DataMap = sl
            c1FlexGrid1.Cols[21].DataMap = sl;

            //事業部
            //SQL変数
            discript = "JigyoubuMei";
            value = "JigyoubuHeadCD";
            table = "Mst_Jigyoubu";
            where = "";
            //where = "JigyoubuHeadCD IS NOT NULL ORDER BY JigyoubuNarabijun";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            DataRow dr;
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }
            src_5.DataSource = combodt;
            src_5.DisplayMember = "Discript";
            src_5.ValueMember = "Value";

            //案件区分
            //SQL変数
            discript = "SakuseiKubun";
            value = "SakuseiKubunID";
            table = "Mst_SakuseiKubun";
            where = "SakuseiKubunID != '03' and SakuseiKubunID != '05' ";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }
            src_6.DataSource = combodt;
            src_6.DisplayMember = "Discript";
            src_6.ValueMember = "Value";

            //SQL変数
            discript = "SakuseiKubun";
            value = "SakuseiKubunID";
            table = "Mst_SakuseiKubun";
            where = "";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }
            //グリッドのコンボボックス用リスト
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとSakuseiKubunをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            c1FlexGrid1.Cols[19].DataMap = sl;
            //c1FlexGrid1.Cols[18].DataMap = sl;

            //契約区分
            //SQL変数
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
            src_9.DataSource = combodt;
            src_9.DisplayMember = "Discript";
            src_9.ValueMember = "Value";
            //グリッドのコンボボックス用リスト
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとSakuseiKubunをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[24].DataMap = sl;
            c1FlexGrid1.Cols[25].DataMap = sl;

            //入札状況
            discript = "RakusatsuShaMei";
            value = "RakusatsuShaID";
            table = "Mst_RakusatsuSha";
            where = "RakusatsuShaNarabijun > 0 ";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }
            src_10.DataSource = combodt;
            src_10.DisplayMember = "Discript";
            src_10.ValueMember = "Value";
            //グリッドのコンボボックス用リスト
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとSakuseiKubunをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[11].DataMap = sl;
            c1FlexGrid1.Cols[12].DataMap = sl;

            //落札者
            discript = "KyougouMeishou";
            value = "KyougouTashaID";
            table = "Mst_KyougouTasha";
            where = "";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            //グリッドのコンボボックス用リスト
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとSakuseiKubunをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[12].DataMap = sl;
            c1FlexGrid1.Cols[13].DataMap = sl;

            //発注者区分1
            discript = "HachushaKubun1Mei";
            value = "HachushaKubun1CD";
            table = "Mst_HachushaKubun1";
            where = "";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }
            src_12.DataSource = combodt;
            src_12.DisplayMember = "Discript";
            src_12.ValueMember = "Value";
            //グリッドのコンボボックス用リスト
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとSakuseiKubunをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[22].DataMap = sl;
            c1FlexGrid1.Cols[23].DataMap = sl;

            //発注者区分2
            discript = "HachushaKubun2Mei";
            value = "HachushaKubun2CD";
            table = "Mst_HachushaKubun2";
            where = "";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            //グリッドのコンボボックス用リスト
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとSakuseiKubunをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[23].DataMap = sl;
            c1FlexGrid1.Cols[24].DataMap = sl;

            //引合状況
            System.Data.DataTable tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));

            tmpdt.Rows.Add(1, "未確定");
            tmpdt.Rows.Add(2, "発注確定");
            tmpdt.Rows.Add(3, "発注無し");
            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }

            src_22.DataSource = tmpdt;
            src_22.DisplayMember = "Discript";
            src_22.ValueMember = "Value";

            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[19].DataMap = sl;
            c1FlexGrid1.Cols[20].DataMap = sl;

            //参考見積
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "未提出");
            tmpdt.Rows.Add(2, "提出");
            tmpdt.Rows.Add(3, "依頼無し");
            tmpdt.Rows.Add(4, "辞退");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[26].DataMap = sl;
            c1FlexGrid1.Cols[27].DataMap = sl;
            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }

            src_20.DataSource = tmpdt;
            src_20.DisplayMember = "Discript";
            src_20.ValueMember = "Value";

            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);


            //受注意欲
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "フラット");
            tmpdt.Rows.Add(2, "あり");
            tmpdt.Rows.Add(3, "なし");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[27].DataMap = sl;
            c1FlexGrid1.Cols[28].DataMap = sl;
            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }

            src_21.DataSource = tmpdt;
            src_21.DisplayMember = "Discript";
            src_21.ValueMember = "Value";


            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "対応前");
            tmpdt.Rows.Add(2, "応札");
            tmpdt.Rows.Add(3, "不参加");
            tmpdt.Rows.Add(4, "辞退");


            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }

            src_23.DataSource = tmpdt;
            src_23.DisplayMember = "Discript";
            src_23.ValueMember = "Value";

            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[28].DataMap = sl;
            c1FlexGrid1.Cols[29].DataMap = sl;

            // 入札方式
            discript = "KeiyakuKeitai";
            value = "KeiyakuKeitaiCD";
            table = "Mst_KeiyakuKeitai";
            //where = "";
            where = "KeiyakuKeitaiNarabijun < 20 ";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            dr = combodt.NewRow();
            combodt.Rows.InsertAt(dr, 0);
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[25].DataMap = sl;
            c1FlexGrid1.Cols[26].DataMap = sl;


            //帳票
            discript = "PrintName";
            value = "PrintListID";
            table = "Mst_PrintList";
            //where = "";
            where = "MENU_ID = 100 AND PrintBunruiCD = 1 AND PrintDelFlg <> 1 ORDER BY PrintListNarabijun";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            //dr = combodt.NewRow();
            //combodt.Rows.InsertAt(dr, 0);
            comboBox13.DataSource = combodt;
            comboBox13.DisplayMember = "Discript";
            comboBox13.ValueMember = "Value";

            //契約図書の画像切り替え
            Hashtable imgMap = new Hashtable();
            imgMap.Add(0, Image.FromFile("Resource/Image/folder_gray_s.png"));
            imgMap.Add(1, Image.FromFile("Resource/Image/folder_yellow_s.png"));
            //c1FlexGrid1.Cols[14].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            //c1FlexGrid1.Cols[14].ImageMap = imgMap;
            //c1FlexGrid1.Cols[14].ImageAndText = false;
            c1FlexGrid1.Cols[15].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[15].ImageMap = imgMap;
            c1FlexGrid1.Cols[15].ImageAndText = false;

            //編集の画像切り替え
            imgMap = new Hashtable();
            imgMap.Add(0, Image.FromFile("Resource/Image/file_presentation1_g.png"));
            imgMap.Add(1, Image.FromFile("Resource/Image/file_presentation1.png"));
            c1FlexGrid1.Cols[2].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[2].ImageMap = imgMap;
            c1FlexGrid1.Cols[2].ImageAndText = false;
        }

        private void set_combo_shibu(string nendo)
        {
            //受託課所支部
            string SelectedValue = "";
            if (src_4.Text != "")
            {
                SelectedValue = src_4.SelectedValue.ToString();
            }
            //SQL変数
            string discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
            string value = "Mst_Busho.GyoumuBushoCD ";
            string table = "Mst_Busho";
            string where = "GyoumuBushoCD < '999990' AND BushoNewOld <= 1 AND BushoEntryHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 " +
                    //"AND NOT GyoumuBushoCD = '150210' AND NOT GyoumuBushoCD = '150410' AND NOT GyoumuBushoCD = '150200' AND NOT GyoumuBushoCD = '150400' " +
                    "AND NOT GyoumuBushoCD LIKE '121%' ";
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
            Console.WriteLine(where);
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            DataRow dr = combodt.NewRow();
            combodt.Rows.InsertAt(dr, 0);
            src_4.DataSource = combodt;
            src_4.DisplayMember = "Discript";
            src_4.ValueMember = "Value";
            if (SelectedValue != "")
            {
                src_4.SelectedValue = SelectedValue;
            }
        }

        private void get_date()
        {
            String nendo1 = "";
            String nendo2 = "";

            String koukinendo1 = "";
            String koukinendo2 = "";

            int year = 0;

            // 年度計算
            // 売上年度
            if (src_1.Text != "") { 
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
            // 工期年度
            if (item1_KoukiNendo.Text != "")
            {
                koukinendo1 = item1_KoukiNendo.Text.Substring(0, 4);
                koukinendo2 = koukinendo1;
                if (item1_2_SanNen.Checked)
                {
                    koukinendo2 = (int.Parse(koukinendo1) - 2).ToString();
                }
            }
            else
            {
                year = int.Parse(GlobalMethod.GetTodayNendo());
                koukinendo1 = year.ToString();
                year = year - 4;
                koukinendo2 = year.ToString();
            }
            
            try
            {
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    string tokai = "";
                    string keichokai = "";
                    string tokaiID = "";
                    string keichokaiID = "";
                    var cmd = conn.CreateCommand();
                    // 当会（建設物価調査会）のIDを取得
                    tokaiID = GlobalMethod.GetCommonValue1("KYOUGOU_TOKAI_ID");
                    if (tokaiID == "" || tokaiID == null)
                    {
                        tokaiID = "";
                        tokai = "建設物価調査会";
                    }
                    else
                    {
                        cmd.CommandText = "SELECT " +
                            "KyougouMeishou " +
                            "FROM Mst_KyougouTasha " +
                            "WHERE KyougouTashaID = '" + tokaiID + "' ";
                        var sda1 = new SqlDataAdapter(cmd);
                        var dt1 = new DataTable();
                        sda1.Fill(dt1);

                        tokai = dt1.Rows[0][0].ToString();
                        if (tokai == "" || tokai == null)
                        {
                            tokaiID = "";
                            tokai = "建設物価調査会";
                        }
                    }
                    keichokaiID = GlobalMethod.GetCommonValue1("KYOUGOU_KEICHO_ID");
                    if (keichokaiID == "" || keichokaiID == null)
                    {
                        keichokaiID = "";
                        keichokai = "（一財）経済調査会";
                    }
                    else
                    {
                        cmd.CommandText = "SELECT " +
                            "KyougouMeishou " +
                            "FROM Mst_KyougouTasha " +
                            "WHERE KyougouTashaID = '" + keichokaiID + "' ";
                        var sda2 = new SqlDataAdapter(cmd);
                        var dt2 = new DataTable();
                        sda2.Fill(dt2);

                        keichokai = dt2.Rows[0][0].ToString();
                        if (keichokai == "" || keichokai == null)
                        {
                            keichokaiID = "";
                            keichokai = "（一財）経済調査会";
                        }
                    }

                    cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT " +
                            "AnkenJouhou.AnkenJouhouID" +
                            ",AnkenKoukiNendo " +
                            ",AnkenUriageNendo " +
                            ",AnkenKeikakuBangou " +
                            ",AnkenAnkenBangou " +
                            ",AnkenJutakuBangou " +
                            ",AnkenHachuushaKaMei " +
                            ",AnkenGyoumuMei " +
                            ",CASE AnkenUriagebi WHEN '1753/01/01' THEN null ELSE FORMAT(AnkenUriagebi,'yyyy/MM') END " +
                            ",NyuusatsuJouhou.NyuusatsuRakusatsushaID " +
                            ",NyuusatsuRakusatsusha " +
                            ",CASE AnkenTourokubi WHEN '1753/01/01' THEN null ELSE FORMAT(AnkenTourokubi,'yyyy/MM/dd') END " +
                            ",CASE AnkenKeiyakusho WHEN null THEN 0 WHEN '' THEN 0 ELSE 1 END " +
                            ",AnkenKeiyakusho " +
                            ",AnkenKianZumi " +
                            ",CASE AnkenKeiyakuSakuseibi WHEN '1753/01/01' THEN null ELSE FORMAT(AnkenKeiyakuSakuseibi,'yyyy/MM/dd') END " +
                            ",AnkenSakuseiKubun " +
                            ",AnkenHikiaijhokyo " +
                            ",AnkenJutakubushoCD " +
                            ",CASE AnkenNyuusatsuYoteibi WHEN '1753/01/01' THEN null ELSE FORMAT(AnkenNyuusatsuYoteibi,'yyyy/MM/dd') END " +
                            ",HachushaKubun1CD " +
                            ",HachushaKubun2CD " +
                            ",AnkenGyoumuKubunMei " +
                            //",NyuusatsuJouhou.NyuusatsuGyoumuBikou " +
                            ",AnkenNyuusatsuHoushiki " + // 入札方式
                            ",AnkenToukaiSankouMitsumori " +
                            ",AnkenToukaiJyutyuIyoku " +
                            ",AnkenToukaiOusatu " +
                            ",ISNULL(NyuusatsuRakusatugaku,0) " +      // 落札額（税抜）
                            ",ISNULL(NyuusatsuYoteiKakaku,0) " +       // 予定価格（税抜）
                            ",ISNULL(KeiyakuHaibunChoZeinuki,0) " +    // 調査部配分額(税抜)

                            ",ISNULL(KeiyakuUriageHaibunCho,0) " +     // 調査部配分額(税込)
                            //",ISNULL(AnkenKeiyakuUriageHaibunGakuC,0) " +     // 調査部配分額(税込)

                            ",ISNULL(KeiyakuTankeiMikomiCho,0) " +     // 調査部単契等の見込補正額(税抜)
                            
                            //",ISNULL(AnkenKeiyakuZeikomiKingaku,0) " +
                            ",ISNULL(KeiyakuZeikomiKingaku,0) " + // 契約金額（税込）

                            ",ISNULL(NyuusatsuMitsumorigaku,0) " +
                            ",NyuusatsuKekkaMemo " +
                            ",ISNULL(KeiyakuHaibunJoZeinuki,0) " +     // 事業普及部配分額(税抜)
                            
                            ",ISNULL(KeiyakuUriageHaibunJo,0) " +      // 事業普及部配分額(税込)
                            //",ISNULL(AnkenKeiyakuUriageHaibunGakuJ,0) " +      // 事業普及部配分額(税込)

                            ",ISNULL(KeiyakuTankeiMikomiJo,0) " +      // 事業普及部単契等の
                            ",ISNULL(KeiyakuHaibunJosysZeinuki,0) " +  // 情報システム配分額(税抜)

                            ",ISNULL(KeiyakuUriageHaibunJosys,0) " +   // 情報システム配分額(税込)
                            //",ISNULL(AnkenKeiyakuUriageHaibunGakuJs,0) " +   // 情報システム配分額(税込)

                            ",ISNULL(KeiyakuTankeiMikomiJosys,0) " +   // 情報システム単契等の
                            ",ISNULL(KeiyakuHaibunKeiZeinuki,0) " +    // 総合研究所配分額(税抜)
                            
                            ",ISNULL(KeiyakuUriageHaibunKei,0) " +     // 総合研究所配分額(税込)
                            //",ISNULL(AnkenKeiyakuUriageHaibunGakuK,0) " +     // 総合研究所配分額(税込)

                            ",ISNULL(KeiyakuTankeiMikomiKei,0) " +     // 総合研究所単契等の
                            ",AnkenTantoushaMei " +
                            ",CASE AnkenKeiyakuKoukiKaishibi WHEN '1753/01/01' THEN null ELSE FORMAT(AnkenKeiyakuKoukiKaishibi,'yyyy/MM/dd') END " +
                            ",CASE AnkenKeiyakuKoukiKanryoubi WHEN '1753/01/01' THEN null ELSE FORMAT(AnkenKeiyakuKoukiKanryoubi,'yyyy/MM/dd') END " +
                            ",CASE AnkenKeiyakuTeiketsubi WHEN '1753/01/01' THEN null ELSE FORMAT(AnkenKeiyakuTeiketsubi,'yyyy/MM/dd') END " +
                            ",CASE WHEN ISNULL(KeiyakuKurikoshiCho,0) > 0 THEN '有' " +
                            "      WHEN ISNULL(KeiyakuKurikoshiJo,0) > 0 THEN '有' " +
                            "      WHEN ISNULL(KeiyakuKurikoshiJosys,0) > 0 THEN '有' " +
                            "      WHEN ISNULL(KeiyakuKurikoshiKei,0) > 0 THEN '有' " +
                            "      ELSE '無' END AS 'NendoKurikoshi' " +
                            ",AnkenHachushaCD " +
                            ",AnkenSaishinFlg " +
                            "FROM AnkenJouhou " +
                            "LEFT JOIN NyuusatsuJouhou ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID " +
                            "LEFT JOIN KeiyakuJouhouEntory ON KeiyakuJouhouEntory.AnkenJouhouID = AnkenJouhou.AnkenJouhouID " +
                            "LEFT JOIN Mst_SakuseiKubun ON AnkenSakuseiKubun = SakuseiKubunID " +
                            "LEFT JOIN Mst_Busho ON AnkenJutakubushoCD = GyoumuBushoCD " +
                            "LEFT JOIN Mst_KeiyakuKeitai ON AnkenNyuusatsuHoushiki = KeiyakuKeitaiCD " +
                            "LEFT JOIN Mst_Hachusha ON AnkenHachushaCD = HachushaCD " +
                            "LEFT JOIN NyuusatsuJouhouOusatsusha ON NyuusatsuJouhouOusatsusha.NyuusatsuJouhouID = NyuusatsuJouhou.NyuusatsuJouhouID AND NyuusatsuRakusatsuJokyou = 1 " +
                            "LEFT JOIN KeikakuJouhou ON KeikakuBangou = AnkenKeikakuBangou " +
                            "LEFT JOIN GyoumuHaibun ON GyoumuHaibun.GyoumuAnkenJouhouID = AnkenJouhou.AnkenJouhouID and GyoumuHaibun.GyoumuHibunKubun = '10'" +
                            "WHERE AnkenUriageNendo <= '" + nendo1 + "' and AnkenUriageNendo >= '" + nendo2 + "' " +
                            "AND AnkenKoukiNendo <= '" + koukinendo1 + "' and AnkenKoukiNendo >= '" + koukinendo2 + "' " + 
                            "AND AnkenDeleteFlag = 0 ";

                    // 受託課所支部が空ではない場合
                    if (src_4.Text != "")
                    {
                        if (src_4.SelectedValue.ToString() != "127100")
                        {
                            // 128400 情報システム部関連で、2021年度よりも前の場合
                            if (src_4.SelectedValue.ToString().Substring(0, 4) == "1284" && item1_KoukiNendo.Text != null && item1_KoukiNendo.Text != "" && int.Parse(item1_KoukiNendo.SelectedValue.ToString()) < 2021)
                            {
                                // 127900 情報システム部【旧】を見る
                                cmd.CommandText += "  and AnkenJutakubushoCD LIKE '127900'";
                            }
                            else if (src_4.SelectedValue.ToString().Substring(0, 4) == "1292")
                            {
                                cmd.CommandText += "  and AnkenJutakubushoCD LIKE '129230'";
                            }
                            else
                            {
                                cmd.CommandText += "  and AnkenJutakubushoCD LIKE '" + src_4.SelectedValue.ToString().TrimEnd('0') + "%'";
                            }

                            if (src_4.SelectedValue.ToString() == "127000")
                            {
                                // 2021年度より前は、1279は除外する
                                //if (int.Parse(src_1.SelectedValue.ToString()) < 2021) {
                                if (item1_KoukiNendo.Text != null && item1_KoukiNendo.Text != "" && int.Parse(item1_KoukiNendo.SelectedValue.ToString()) < 2021)
                                {
                                        // 127000 本部 調査部門の場合、 1279 情報システム部関連は除外
                                        cmd.CommandText += "  and NOT AnkenJutakubushoCD LIKE '1279%'";
                                }
                            }
                        }
                        else
                        {
                            cmd.CommandText += "  and AnkenJutakubushoCD like '1271%' ";
                        }
                    }

                    if (src_5.Text == "調査部")
                    {
                        //cmd.CommandText += "  and ((KeiyakuJouhouEntory.KeiyakuUriageHaibunGakuC <> 0 ) or (AnkenGyoumuKubunCD <> '04' AND AnkenGyoumuKubunCD <> '05' AND AnkenGyoumuKubunCD <> '11' AND AnkenGyoumuKubunCD <> '12' AND AnkenGyoumuKubunCD <> '13')) ";
                        cmd.CommandText += "  and ((AnkenKeiyakuUriageHaibunGakuC <> 0 ) or (AnkenGyoumuKubunCD <> '04' AND AnkenGyoumuKubunCD <> '05' AND AnkenGyoumuKubunCD <> '11' AND AnkenGyoumuKubunCD <> '12' AND AnkenGyoumuKubunCD <> '13')) ";
                    }
                    else if (src_5.Text == "事業普及部")
                    {
                        //cmd.CommandText += "  and ((KeiyakuJouhouEntory.KeiyakuUriageHaibunGakuJ <> 0 ) or (KeiyakuJouhouEntory.KeiyakuUriageHaibunGakuR <> 0 ) or (AnkenGyoumuKubunCD = '12' or AnkenGyoumuKubunCD = '13')) ";
                        cmd.CommandText += "  and ((AnkenKeiyakuUriageHaibunGakuJ <> 0 ) or (AnkenGyoumuKubunCD = '12' or AnkenGyoumuKubunCD = '13')) ";
                    }
                    else if (src_5.Text == "情報システム部")
                    {
                        //cmd.CommandText += "  and ((KeiyakuJouhouEntory.KeiyakuUriageHaibunGakuJs <> 0 ) or (AnkenGyoumuKubunCD = '04' OR AnkenGyoumuKubunCD = '11' )) ";
                        cmd.CommandText += "  and ((AnkenKeiyakuUriageHaibunGakuJs <> 0 ) or (AnkenGyoumuKubunCD = '04' OR AnkenGyoumuKubunCD = '11' )) ";
                    }
                    else if (src_5.Text == "総合研究所")
                    {
                        //cmd.CommandText += "  and ((KeiyakuJouhouEntory.KeiyakuUriageHaibunGakuK <> 0 ) or (AnkenGyoumuKubunCD = '05')) ";
                        cmd.CommandText += "  and ((AnkenKeiyakuUriageHaibunGakuK <> 0 ) or (AnkenGyoumuKubunCD = '05')) ";
                    }
                    if (src_6.Text != "")
                    {
                        cmd.CommandText += "  and AnkenSakuseiKubun = '" + src_6.SelectedValue + "'";
                    }
                    if (src_7.CustomFormat == "")
                    {
                        cmd.CommandText += "  and AnkenNyuusatsuYoteibi >= '" + src_7.Text + "'";
                    }
                    if (src_8.CustomFormat == "")
                    {
                        cmd.CommandText += "  and AnkenNyuusatsuYoteibi <= '" + src_8.Text + "'";
                    }
                    if (src_9.Text != "")
                    {
                        cmd.CommandText += "  and AnkenGyoumuKubun = " + src_9.SelectedValue + "";
                    }
                    if (src_10.Text != "")
                    {
                        cmd.CommandText += "  and NyuusatsuRakusatsushaID = " + src_10.SelectedValue + "";
                    }
                    if (src_11.Text != "")
                    {
                        cmd.CommandText += "  and NyuusatsuRakusatsusha COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_11.Text, 1, 0) + "%' ESCAPE '\\' ";
                    }

                    if (src_12.Text != "")
                    {
                        cmd.CommandText += "  and HachushaKubun1CD = '" + src_12.SelectedValue + "'";
                    }
                    if (src_13.Text != "")
                    {
                        cmd.CommandText += "  and AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_13.Text, 1, 0) + "%' ESCAPE '\\' ";
                    }
                    if (src_14.Text != "")
                    {
                        //cmd.CommandText += "  and AnkenKeikakuAnkenMei like '%" + GlobalMethod.ChangeSqlText(src_14.Text, 1, 0) + "%' ESCAPE '\\' ";
                        cmd.CommandText += "  and KeikakuAnkenMei COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_14.Text, 1, 0) + "%' ESCAPE '\\' ";
                    }

                    if (src_15.Text != "")
                    {
                        cmd.CommandText += "  and AnkenJutakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_15.Text, 1, 0) + "%' ESCAPE '\\' ";
                    }
                    if (src_16.Text != "")
                    {
                        cmd.CommandText += "  and AnkenAnkenBangou COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_16.Text, 1, 0) + "%' ESCAPE '\\' ";
                    }
                    if (src_17.Text != "")
                    {
                        cmd.CommandText += "  and AnkenGyoumuMei COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_17.Text, 1, 0) + "%' ESCAPE '\\' ";
                    }
                    if (src_18.Text != "")
                    {
                        cmd.CommandText += "  and AnkenHachuushaKaMei COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_18.Text, 1, 0) + "%' ESCAPE '\\' ";
                    }
                    if (src_19.Text != "")
                    {
                        cmd.CommandText += "  and CONVERT(NVARCHAR, AnkenUriagebi, 111) COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_19.Text, 1, 0) + "%' ESCAPE '\\' ";
                    }

                    if (src_20.Text != "")
                    {
                        cmd.CommandText += "  and AnkenToukaiSankouMitsumori = N'" + src_20.SelectedIndex + "'";
                    }
                    if (src_21.Text != "")
                    {
                        cmd.CommandText += "  and AnkenToukaiJyutyuIyoku = N'" + src_21.SelectedIndex + "'";
                    }
                    if (src_22.Text != "")
                    {
                        cmd.CommandText += "  and AnkenHikiaijhokyo = N'" + src_22.SelectedIndex + "'";
                    }
                    if (src_23.Text != "")
                    {
                        cmd.CommandText += "  and AnkenToukaiOusatu = N'" + src_23.SelectedIndex + "'";
                    }
                    if (src_24.Checked == true)
                    {
                        cmd.CommandText += "  and ( ISNULL(KeiyakuKurikoshiCho,0) > 0 or ISNULL(KeiyakuKurikoshiJo,0) > 0 or ISNULL(KeiyakuKurikoshiJosys,0) > 0 or ISNULL(KeiyakuKurikoshiKei,0) > 0 ) ";
                    }
                    if (src_25.Text != "")
                    {
                        //cmd.CommandText += "  and AnkenHachuushaCodeID like '%" + GlobalMethod.ChangeSqlText(src_25.Text, 1, 0) + "%' ESCAPE '\\' ";
                        cmd.CommandText += "  and AnkenHachushaCD COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(src_25.Text, 1, 0) + "%' ESCAPE '\\' ";
                    }
                    if (src_26.Text == "未")
                    {
                        cmd.CommandText += "  and ISNULL(AnkenKianZumi,0) = 0";
                    }
                    else if (src_26.Text == "済")
                    {
                        cmd.CommandText += "  and AnkenKianZumi = 1";
                    }
                    if (src_27.Checked == true)
                    {
                        cmd.CommandText += "  and AnkenSaishinFlg = 1";
                    }

                    // 499 業務日報用の案件を除外する
                    //cmd.CommandText += " and AnkenAnkenBangou not like '%999' ";
                    // 業務日報のデータは60000～70000の間で登録している為、除外
                    cmd.CommandText += " and (AnkenJouhou.AnkenJouhouID < 60000 or AnkenJouhou.AnkenJouhouID > 70000) ";

                    //初期ソート処理予定
                    cmd.CommandText += " ORDER BY " +
                    " CASE " +
                    " WHEN NyuusatsuRakusatsushaID = 1 THEN '000' " +
                    " WHEN NyuusatsuRakusatsushaID = 2 AND AnkenKianZumi <> 1 AND ISNULL(NyuusatsuRakusatsusha,'') = '' THEN '010' " +
                    " WHEN NyuusatsuRakusatsushaID = 2 AND AnkenKianZumi <> 1 AND NyuusatsuRakusatsusha COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + tokai + "' THEN '020' " +
                    " WHEN NyuusatsuRakusatsushaID = 2 AND AnkenKianZumi <> 1 AND NyuusatsuRakusatsusha COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + keichokai + "' THEN '030' " +
                    " WHEN NyuusatsuRakusatsushaID = 2 AND AnkenKianZumi <> 1 THEN '040' " +
                    " WHEN NyuusatsuRakusatsushaID = 2 AND AnkenKianZumi = 1 THEN '050' " +
                    " WHEN NyuusatsuRakusatsushaID = 3 THEN '060' " +
                    " WHEN NyuusatsuRakusatsushaID = 4 THEN '070' " +
                    " ELSE NyuusatsuRakusatsushaID " +
                    " END , AnkenKoukiNendo, AnkenUriageNendo, AnkenJutakuBangou, AnkenJutakuBangouEda, AnkenJutakubushoCD, NyuusatsuRakusatsusha desc, AnkenJouhouID, AnkenHachushaKubunCD, AnkenGyoumuMei";


                    GlobalMethod.outputLogger("Search_Entry", "開始", "GetAnkenJouhou", UserInfos[1]);
                    var sda = new SqlDataAdapter(cmd);
                    if (src_4.Text == "")
                    {
                        // 全国検索の為データ件数が多くなりますが、よろしいですか？ 20210225 聞かないようにする
                        //if (GlobalMethod.outputMessage("I10501", "") == DialogResult.OK)
                        //{
                            ListData.Clear();
                            sda.Fill(ListData);
                        //}
                    }
                    else
                    {
                        ListData.Clear();
                        sda.Fill(ListData);
                    }
                    if (ListData.Rows.Count == 0)
                    {
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("I20001", ""));
                    }
                }
                Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / int.Parse(src_28.Text.Replace("件", "")))).ToString();
                Paging_now.Text = (1).ToString();
                set_data(int.Parse(Paging_now.Text));
                set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
                GlobalMethod.outputLogger("Search_Entry", "終了", "GetAnkenJouhou", UserInfos[1]);

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
            catch (Exception)
            {

                throw;
            }
        }

        private void set_data(int pagenum)
        {
            c1FlexGrid1.Visible = false;

            c1FlexGrid1.Rows.Count = 2;
            c1FlexGrid1.AllowAddNew = true;
            int viewnum = int.Parse(src_28.Text.Replace("件", "")) + 2;
            int startrow = (pagenum - 1) * viewnum;
            int addnum = ListData.Rows.Count - startrow;
            if (addnum > viewnum)
            {
                addnum = viewnum;
            }
            for (int r = 0; r < addnum; r++)
            {
                c1FlexGrid1.Rows.Add();
                if (UserInfos[1].Substring(0, 3) == "121")
                {
                    c1FlexGrid1[r + 2, 2] = 0;
                }
                else
                {
                    c1FlexGrid1[r + 2, 2] = 1;
                }
                for (int i = 0; i < c1FlexGrid1.Cols.Count - 3; i++)
                {
                    if (i == 12)
                    {
                        if (Directory.Exists(ListData.Rows[startrow + r][i + 1].ToString()))
                        {
                            //c1FlexGrid1[r + 2, i + 3] = 1;
                            c1FlexGrid1[r + 2, i + 3] = 1;
                        }
                        else
                        {
                            //c1FlexGrid1[r + 2, i + 3] = 0;
                            c1FlexGrid1[r + 2, i + 3] = 0;
                        }
                    }
                    else
                    {
                        c1FlexGrid1[r + 2, i + 3] = ListData.Rows[startrow + r][i];
                    }
                }
                c1FlexGrid1.Rows[r + 2].Height = 40;
            }
            c1FlexGrid1.AllowAddNew = false;
            if (c1FlexGrid1.Rows.Count > 2)
            {
                /*
                C1.Win.C1FlexGrid.CellRange cr;
                cr = c1FlexGrid1.GetCellRange(2, 2, c1FlexGrid1.Rows.Count - 1, 2);
                cr.Image = Image.FromFile("Resource/file_presentation1.png");
                cr = c1FlexGrid1.GetCellRange(2, 10, c1FlexGrid1.Rows.Count - 1, 10);
                cr.Style = c1FlexGrid1.Styles["CelStyle"];
                */
                c1FlexGrid1.Select(2, 2, true);
            }

            c1FlexGrid1.Visible = true;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
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

        private void pictureBox3_Click(object sender, EventArgs e)
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

        private void pictureBox4_Click(object sender, EventArgs e)
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

        private void pictureBox5_Click(object sender, EventArgs e)
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

        private void button1_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            set_error("", 0);
            // エラーフラグ true：正常 false：エラー
            Boolean errorflg = false;
            errorflg = uriagebi_Check();

            if (errorflg != false) {
                //set_error("", 0);
                get_date();
            }

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private Boolean uriagebi_Check()
        {
            // エラーフラグ true：正常 false：エラー
            Boolean errorflg = true;
            // 管理月チェック
            if (src_19.Text != "" && src_19.Text != null)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(src_19.Text, @"^\d{4}/\d{2}$"))
                {
                    int year;
                    int month;

                    // YYYY/MM 形式チェック
                    if (int.TryParse(src_19.Text.Substring(0, 4), out year) && int.TryParse(src_19.Text.Substring(5, 2), out month))
                    {
                        if ((year < 1900 || year > 2100) || (month < 1 || month > 12))
                        {
                            set_error(GlobalMethod.GetMessage("E10004", ""));
                            errorflg = false;
                        }
                    }
                    else
                    {
                        set_error(GlobalMethod.GetMessage("E10004", ""));
                        errorflg = false;
                    }
                }
                else
                {
                    set_error(GlobalMethod.GetMessage("E10005", ""));
                    errorflg = false;
                }
            }
            return errorflg;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            ClearForm();

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void ClearForm()
        {
            //検索条件初期化
            //売上年度　受託課所支部
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
            }
            else
            {
                src_1.SelectedValue = System.DateTime.Now.Year;
            }
            */
            src_1.SelectedValue = GlobalMethod.GetTodayNendo();
            //set_combo_shibu(src_1.SelectedValue.ToString());

            item1_KoukiNendo.SelectedValue = GlobalMethod.GetTodayNendo();
            set_combo_shibu(item1_KoukiNendo.SelectedValue.ToString());

            item1_1_Tounendo.Checked = false;
            item1_2_SanNen.Checked = true;

            src_2.Checked = true;
            src_3.Checked = false;
            src_4.SelectedValue = UserInfos[2];
            //事業部　案件区部
            src_5.SelectedIndex = -1;
            src_6.SelectedIndex = -1;
            //入札予定日　from end
            src_7.CustomFormat = " ";
            src_8.CustomFormat = " ";
            //契約区分　入札状況　発注者区分
            src_9.SelectedIndex = -1;
            src_10.SelectedIndex = -1;
            src_12.SelectedIndex = -1;
            //落札者　計画番号　計画案件名
            src_11.Text = "";
            src_13.Text = "";
            src_14.Text = "";
            //受託番号　案件番号　業務名称　発注者名・課名　管理月
            src_15.Text = "";
            src_16.Text = "";
            src_17.Text = "";
            src_18.Text = "";
            src_19.Text = "";
            //参考見積　受注意欲　引合状況　当会応礼
            src_20.SelectedIndex = -1;
            src_21.SelectedIndex = -1;
            src_22.SelectedIndex = -1;
            src_23.SelectedIndex = -1;
            //年度越え配分　発注者CD
            src_24.Checked = false;
            src_25.Text = "";
            //起票状況　最新伝票　表示件数
            src_26.SelectedIndex = 0;
            src_27.Checked = true;
            src_28.SelectedIndex = 1;

            //グリッドコントロールを初期化
            c1FlexGrid1.Styles.Normal.WordWrap = true;
            c1FlexGrid1.Rows[0].AllowMerging = true;
            c1FlexGrid1.AllowAddNew = false;


            if (c1FlexGrid1.Rows.Count > 2)
            {
                //グリッドクリア ヘッダー以外削除
                c1FlexGrid1.Rows.Count = 2;
            }
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
        }

        // ソート前処理
        private void c1FlexGrid1_BeforeSort(object sender, C1.Win.C1FlexGrid.SortColEventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            // データ取り直す
            get_date();
            DataView dv = new DataView(ListData);
            dv.Sort = ListData.Columns[e.Col - 3].ColumnName;
            if (c1FlexGrid1.Cols[e.Col].Sort == C1.Win.C1FlexGrid.SortFlags.Ascending)
            {
                dv.Sort += " DESC";
            }
            ListData = dv.ToTable();
            set_data(int.Parse(Paging_now.Text));

            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {

            Popup_Kyougou form = new Popup_Kyougou();

            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                src_11.Text = form.ReturnValue[1];
            }
        }

        private void c1FlexGrid1_BeforeEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            // e.Row > 1                              ：ヘッダー行より下の項目かどうか
            // e.COl == 11                            ：管理月入力×をクリックしたか
            //  UserInfos[2].Length >= 3              ：ログインユーザーの部所CDが3桁以上か
            // (UserInfos[2].Substring(0, 3) == "121" ：ログインユーザーの部所CD頭3桁が121でないか
            // または、
            // !UserInfos[2].Equals(c1FlexGrid1[e.Row, 21].ToString())：ログインユーザーの部所と受託課所支部が異なっている場合
            // 編集をキャンセルする
            //if (e.Row > 1 && e.Col == 10 && UserInfos[2].Length >= 3 && (UserInfos[2].Substring(0, 3) == "121" || !UserInfos[2].Equals(c1FlexGrid1[e.Row, 20].ToString())))
            if (e.Row > 1 && e.Col == 11)
            {
                // ロールによる制限　管理職で他部所の場合、管理月更新不可
                // 管理部門（121:総務、経理は管理月の更新が可能）
                if (UserInfos[2].Length >= 3 && (UserInfos[2].Substring(0, 3) != "121" && "1".Equals(UserInfos[4]) && !UserInfos[2].Equals(c1FlexGrid1[e.Row, 21].ToString())))
                { 
                    e.Cancel = true;
                }
            }
        }


        private void src_1_TextChanged(object sender, EventArgs e)
        {
            //set_combo_shibu(src_1.SelectedValue.ToString());
            set_combo_shibu(item1_KoukiNendo.SelectedValue.ToString());
        }

        private void src_3_Click(object sender, EventArgs e)
        {
            //set_combo_shibu(src_1.SelectedValue.ToString());
            set_combo_shibu(item1_KoukiNendo.SelectedValue.ToString());
        }

        private void src_4_TextChanged(object sender, EventArgs e)
        {
            //if (src_4.Text == "")
            //{
            //    button3.Enabled = false;
            //    button3.BackColor = Color.FromArgb(169, 169, 169);
            //}
            //else
            //{
            //    button3.Enabled = true;
            //    button3.BackColor = Color.FromArgb(42, 78, 122);
            //}
        }

        // 帳票出力ボタン
        private void button5_Click(object sender, EventArgs e)
        {
            string AnkenJouhouID = "";

            if (comboBox13.Text == "")
            {
                set_error("", 0);
                set_error("帳票を選択してください。");
            }
            else
            {
                // Gridで選択されている案件番号を取得する
                for (int i = 2; i < c1FlexGrid1.Rows.Count; i++)
                {
                    if (c1FlexGrid1.GetCellCheck(i, 1) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                    {
                        AnkenJouhouID = c1FlexGrid1[i, 3].ToString();
                        break;
                    }
                }

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
                      "WHERE PrintListID = '" + comboBox13.SelectedValue + "'";

                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(Dt);
                    Boolean errorFLG = false;

                    if (Dt.Rows.Count > 0)
                    {
                        set_error("", 0);
                        // 3:売上明細表の場合、PrintKikanFlg が 1 で、入札（予定）日のFromかToのどちらかでも空の場合、エラー
                        if (Dt.Rows[0][0].ToString() == "3" && Dt.Rows[0][1].ToString() == "1")
                        {
                            // 入札（予定）日From
                            if (src_7.CustomFormat == " ")
                            {
                                errorFLG = true;
                                // 受託番号
                                src_15.BackColor = Color.FromArgb(255, 204, 255);
                            }
                            // 入札（予定）日To
                            if (src_8.CustomFormat == " ")
                            {
                                errorFLG = true;
                                // 受託番号
                                src_15.BackColor = Color.FromArgb(255, 204, 255);
                            }
                        }
                        // 66:管理票の場合、受託番号が空の場合、エラー
                        if (Dt.Rows[0][0].ToString() == "66")
                        {
                            if (src_15.Text == "")
                            {
                                errorFLG = true;
                                // 受託番号
                                src_15.BackColor = Color.FromArgb(255, 204, 255);
                            }
                        }
                        // 75:エントリくん一覧出力(新）
                        if (Dt.Rows[0][0].ToString() == "75")
                        {
                            set_error("", 0);
                            // エラーフラグ true：正常 false：エラー
                            Boolean errorflg = true;
                            // 管理月のFormatをチェック
                            errorflg = uriagebi_Check();

                            if (errorflg != false)
                            {

                                // string[]
                                // 1：UriageNendo                検索条件.売上年度                             src_1  売上年度
                                // 2：UriageNendoOption          当年度の場合、1 3年以内の場合、2              src_2(当年度) or src_3(3年以内)
                                // 3：JutakuKashoshibuCD         検索条件.受託課所支部CD                       src_4  受託課所支部
                                // 4：JigyoubuCD                 検索条件.事業部CD                             src_5  事業部
                                // 5：AnkenKubunCD               検索条件.案件区分CD                           src_6  案件区分
                                // 6：NyuusatsuYoteibiFrom       検索条件.入札予定日from                       src_7  入札（予定）日From
                                // 7：NyuusatsuYoteibiTo         検索条件.入札予定日to                         src_8  入札（予定）日To
                                // 8：KeiyakuKubunCD             検索条件.契約区分CD                           src_9  契約区分
                                // 9：NyusatsuJokyouCD           検索条件.入札状況CD                           src_10 入札状況
                                // 10：Rakusatsusha              検索条件.落札者                               src_11 落札者
                                // 11：HachushaKubun1            検索条件.発注者区分１                         src_12 発注者区分1
                                // 12：KeikakuBangou             検索条件.計画番号                             src_13 計画番号
                                // 13：KeikakuAnkenMei           検索条件.計画案件名                           src_14 計画案件名
                                // 14：JutakuBangou              検索条件.受託番号                             src_15 受託番号
                                // 15：AnkenBangou               検索条件.案件番号                             src_16 案件番号
                                // 16：GyoumuMei                 検索条件.業務名称                             src_17 業務名称
                                // 17：HachuushaKaMei            検索条件.発注者名・課名                       src_18 発注者名・課名
                                // 18：Kanriduki                 検索条件.管理月                               src_19 管理月
                                // 19：SankouMitsumori           検索条件.参考見積CD                           src_20 参考見積
                                // 20：JyutyuIyoku               検索条件.受注意欲CD                           src_21 受注意欲
                                // 21：Hikiaijhokyo              検索条件.引合状況CD                           src_22 引合状況
                                // 22：ToukaiOusatu              検索条件.当会応札CD                           src_23 当会応札
                                // 23：NendogoeHaibun            年度越え配分のチェック有の場合、1無の場合、0  src_24 年度越え配分有
                                // 24：HachuushaCD               検索条件.発注者CD                             src_25 発注者コード
                                // 25：KianJokyo                 検索条件.起案状況CD                           src_26 起案状況
                                // 26：SaishinDenpyou            最新伝票のチェック有の場合、1無の場合、0      src_27 最新伝票
                                // 27：HyouziKensuu              検索条件.表示件数                             src_28 表示件数
                                // 28：KoukiKaishiNendo          検索条件.工期開始年度                         item1_KoukiNendo 工期開始年度
                                // 29：KoukiKaishiNendoOption    当年度の場合、1 3年以内の場合、2              item1_1_Tounendo(当年度) or item1_2_SanNen(3年以内)

                                // 29個分先に用意
                                string[] report_data = new string[29] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                                report_data[0] = src_1.SelectedValue.ToString();
                                // 売上年度オプション
                                if (src_2.Checked)
                                {
                                    report_data[1] = "1";
                                }
                                else
                                {
                                    report_data[1] = "2";
                                }
                                report_data[2] = src_4.SelectedValue.ToString();   // 受託課所支部
                                if (src_5.Text != null && src_5.Text != "")
                                {
                                    report_data[3] = src_5.SelectedValue.ToString();   // 事業部
                                }
                                if (src_6.Text != null && src_6.Text != "")
                                {
                                    report_data[4] = src_6.SelectedValue.ToString();   // 案件区分
                                }
                                if (src_7.CustomFormat == "")
                                {
                                    report_data[5] = "'" + src_7.Text + "'";   // 入札（予定）日From
                                }
                                else
                                {
                                    report_data[5] = "null";
                                }
                                if (src_8.CustomFormat == "")
                                {
                                    report_data[6] = "'" + src_8.Text + "'";   // 入札（予定）日To
                                }
                                else
                                {
                                    report_data[6] = "null";
                                }
                                if (src_9.Text != null && src_9.Text != "")
                                {
                                    report_data[7] = src_9.SelectedValue.ToString();   // 契約区分
                                }
                                else
                                {
                                    report_data[7] = "0";
                                }
                                if (src_10.Text != null && src_10.Text != "")
                                {
                                    report_data[8] = src_10.SelectedValue.ToString();  // 入札状況
                                }
                                else
                                {
                                    report_data[8] = "0";
                                }
                                report_data[9] = src_11.Text;  // 落札者
                                if (src_12.Text != null && src_12.Text != "")
                                {
                                    report_data[10] = src_12.SelectedValue.ToString(); // 発注者区分1
                                }
                                else
                                {
                                    report_data[10] = "0";
                                }
                                report_data[11] = src_13.Text; // 計画番号
                                report_data[12] = src_14.Text; // 計画案件名
                                report_data[13] = src_15.Text; // 受託番号
                                report_data[14] = src_16.Text; // 案件番号
                                report_data[15] = src_17.Text; // 業務名称
                                report_data[16] = src_18.Text; // 発注者名・課名
                                if (src_19.Text != null && src_19.Text != "")
                                {
                                    report_data[17] = "'" + src_19.Text + "/1" + "'"; // 管理月
                                }
                                else
                                {
                                    report_data[17] = "null";
                                }

                                if (src_20.Text != null && src_20.Text != "")
                                {
                                    report_data[18] = src_20.SelectedValue.ToString(); // 参考見積
                                }
                                else
                                {
                                    report_data[18] = "0";
                                }
                                if (src_21.Text != null && src_21.Text != "")
                                {
                                    report_data[19] = src_21.SelectedValue.ToString(); // 受注意欲
                                }
                                else
                                {
                                    report_data[19] = "0";
                                }
                                if (src_22.Text != null && src_22.Text != "")
                                {
                                    report_data[20] = src_22.SelectedValue.ToString(); // 引合状況
                                }
                                else
                                {
                                    report_data[20] = "0";
                                }
                                if (src_23.Text != null && src_23.Text != "")
                                {
                                    report_data[21] = src_23.SelectedValue.ToString(); // 当会応札
                                }
                                else
                                {
                                    report_data[21] = "0";
                                }
                                // 年度越え配分有
                                if (src_24.Checked)
                                {
                                    report_data[22] = "1";
                                }
                                else
                                {
                                    report_data[22] = "2";
                                }
                                report_data[23] = src_25.Text; // 発注者コード

                                // 起案状況
                                if (src_26.Text == "未")
                                {
                                    report_data[24] = "0";
                                }
                                else if (src_26.Text == "済")
                                {
                                    report_data[24] = "1";
                                }
                                else
                                {
                                    report_data[24] = "";
                                }

                                if (src_27.Checked)
                                {
                                    report_data[25] = "1";
                                }
                                else
                                {
                                    // SE 20220217 No.1278 設定値の誤りを修正
                                    //report_data[25] = "2";    // CHG 20220217
                                    report_data[25] = "0";      // CHG 20220217
                                }
                                report_data[26] = src_28.Text; // 表示件数
                                report_data[27] = item1_KoukiNendo.SelectedValue.ToString();
                                // 売上年度オプション
                                if (item1_1_Tounendo.Checked)
                                {
                                    report_data[28] = "1";
                                }
                                else
                                {
                                    report_data[28] = "2";
                                }

                                string[] result = GlobalMethod.InsertReportWork(230, UserInfos[0], report_data);

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
                        // 2:契約図書保管チェック表、43:ISO書式集、44:単価契約見積書書式集、45:着手完了届書式集、46:使用印鑑簿
                        if (Dt.Rows[0][0].ToString() == "2"
                            || Dt.Rows[0][0].ToString() == "43"
                            || Dt.Rows[0][0].ToString() == "44"
                            || Dt.Rows[0][0].ToString() == "45"
                            || Dt.Rows[0][0].ToString() == "46"
                            )
                        {
                            if (AnkenJouhouID == "")
                            {
                                // E10002:契約情報を選択してください。
                                set_error(GlobalMethod.GetMessage("E10002", ""));
                            }
                            else
                            {
                                string[] report_data = new string[1] { "" };
                                report_data[0] = AnkenJouhouID;

                                string[] result = GlobalMethod.InsertReportWork(int.Parse(comboBox13.SelectedValue.ToString()), UserInfos[0], report_data);

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
                        // 65:受託実績統括表
                        if (Dt.Rows[0][0].ToString() == "65")
                        {
                            string[] report_data = new string[1] { "" };
                            report_data[0] = src_1.SelectedValue.ToString();

                            string[] result = GlobalMethod.InsertReportWork(int.Parse(comboBox13.SelectedValue.ToString()), UserInfos[0], report_data);

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
                        if (errorFLG == true)
                        {
                            set_error("必須入力項目が入力されていません。");
                        }
                    }
                    conn.Close();
                }
            }
        }

        private void c1FlexGrid1_CellChecked(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            for (int i = 2; i < c1FlexGrid1.Rows.Count; i++)
            {
                if (e.Row != i)
                {
                    c1FlexGrid1.SetCellCheck(i, 1, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
                }
            }
        }

        private void Entry_Search_ResizeEnd(object sender, EventArgs e)
        {
            tableLayoutPanel9.Invalidate();
            tableLayoutPanel5.Invalidate();
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

        // Activeになったときの処理
        private void Entry_Search_Activated(object sender, EventArgs e)
        {
            if (ReSearch)
            {
                get_date();
                ReSearch = false;
            }
        }
        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        private void item1_KoukiNendo_SelectedIndexChanged(object sender, EventArgs e)
        {
            set_combo_shibu(item1_KoukiNendo.SelectedValue.ToString());
        }

        private void btnGridSize_Click(object sender, EventArgs e)
        {
            //if (btnGridSize.Text == "一覧拡大")
            //{
            //    // height:570 → 1086・・・調査品目明細と合わせる
            //    // width:1864 → 3752
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid1.Height = 1086;
            //    c1FlexGrid1.Width = 3752;
            //}
            //else
            //{
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid1.Height = 574;
            //    c1FlexGrid1.Width = 1864;
            //}
            gridSizeChange();
        }
        private void gridSizeChange()
        {
            //if (btnGridSize.Text == "一覧拡大")
            //{
            //    // height:570 → 1086・・・調査品目明細と合わせる
            //    // width:1864 → 3752
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid1.Height = 1086;
            //    c1FlexGrid1.Width = 3752;
            //}
            //else
            //{
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid1.Height = 574;
            //    c1FlexGrid1.Width = 1864;
            //}
            string num = "";
            int bigHeight = 0;
            int bigWidth = 0;
            int smallHeight = 0;
            int smallWidth = 0;

            if (btnGridSize.Text == "一覧拡大")
            {
                num = GlobalMethod.GetCommonValue1("ENTORY_GRID_BIG_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out bigHeight);
                    if (bigHeight == 0)
                    {
                        bigHeight = 1086;
                    }
                }
                num = GlobalMethod.GetCommonValue1("ENTORY_GRID_BIG_WIDTH");
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
                num = GlobalMethod.GetCommonValue1("ENTORY_GRID_SMALL_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out smallHeight);
                    if (smallHeight == 0)
                    {
                        smallHeight = 574;
                    }
                }
                num = GlobalMethod.GetCommonValue1("ENTORY_GRID_SMALL_WIDTH");
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

