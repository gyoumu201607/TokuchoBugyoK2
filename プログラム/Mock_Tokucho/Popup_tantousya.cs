using C1.Win.C1FlexGrid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TokuchoBugyoK2
{
    public partial class Popup_tantousya : Form
    {
        GlobalMethod GlobalMethod = new GlobalMethod();
        public C1FlexGrid data = new C1FlexGrid();
        public C1FlexGrid Tantoudata = new C1FlexGrid();
        public C1FlexGrid Jimusyodata = new C1FlexGrid();
        public C1FlexGrid Output = new C1FlexGrid();
        int pagelimit = 20;
        public string AnkenID = "";
        public string JutakuBangou = "";
        public Popup_tantousya()
        {
            InitializeComponent();
        }

        private void Popup_tantousya_Load(object sender, EventArgs e)
        {
            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid0.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid0.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            
            //IMEモードの設定
            Tel.ImeMode = ImeMode.Disable; //  電話
            Mail.ImeMode = ImeMode.Disable; //  メール
            Fax.ImeMode = ImeMode.Disable; //  FAX
            set_data();
        }
        private void set_data()
        {
            //ヘッダーデータの取得
            //受託番号
            Header_JutakuBangou.Text = JutakuBangou;
            DataTable dt = new DataTable();
            string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                //SQL生成
                cmd.CommandText = "SELECT HachushaMei, AnkenHachushaKaMei " +
                 "FROM AnkenJouhou " +
                 "LEFT JOIN Mst_Hachusha ON AnkenHachushaCD = HachushaCD " +
                 "WHERE AnkenJouhouID = " + AnkenID;
                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
            }
            if (dt != null && dt.Rows.Count > 0)
            {
                Header_HachushaMei.Text = dt.Rows[0][0].ToString();
                Header_HachushaKamei.Text = dt.Rows[0][1].ToString();
            }

            //工事事務所GRIDの初期化
            c1FlexGrid0.Rows.Count = 1;

            //単価契約画面から受け取ったデータの表示
            for (int i = 1; i < Jimusyodata.Rows.Count; i++)
            {
                if (Jimusyodata.Rows[i][2] != null && Jimusyodata.Rows[i][2].ToString() != "")
                {
                    c1FlexGrid0.Rows.Add();
                    for (int k = 2; k < c1FlexGrid1.Cols.Count; k++)
                    {
                        // VIPS 20220414 コンポーネント最新化にあたり修正
                        if (k < 7)
                        {
                           c1FlexGrid0.Rows[i][k] = Jimusyodata.Rows[i][k - 1];
                        }
                    }
                }
            }

            Jimusyodata = null;

            //担当者データの取得
            Tantoudata.Rows.Count = 0;
            Tantoudata.Cols.Count = 0;

            for (int i = 0; i < data.Rows.Count; i++)
            {
                Tantoudata.Rows.Add();
                for (int k = 0; k < data.Cols.Count; k++)
                {
                    if (i == 0)
                    {
                        Tantoudata.Cols.Add();
                    }
                    Tantoudata.Rows[i][k] = data.Rows[i][k];
                }
            }

            data = null;

            Paging_all0.Text = (Math.Ceiling(((double)c1FlexGrid0.Rows.Count - 1) / pagelimit)).ToString();
            Grid_Visible0(int.Parse(Paging_now0.Text));
            Paging_all.Text = (Math.Ceiling(((double)c1FlexGrid1.Rows.Count - 1) / pagelimit)).ToString();
            Grid_Visible(int.Parse(Paging_now.Text));
        }


        private void Top_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (1).ToString();
            Grid_Visible(int.Parse(Paging_now.Text));
        }

        private void Previous_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_now.Text) - 1).ToString();
            Grid_Visible(int.Parse(Paging_now.Text));
        }

        private void After_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_now.Text) + 1).ToString();
            Grid_Visible(int.Parse(Paging_now.Text));
        }

        private void End_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_all.Text)).ToString();
            Grid_Visible(int.Parse(Paging_now.Text));
        }

        private void Grid_Visible(int page)
        {
            for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
            {
                if ((page - 1) * pagelimit < i && i < page * pagelimit + 1)
                {
                    c1FlexGrid1.Rows[i].Visible = true;
                }
                else
                {
                    c1FlexGrid1.Rows[i].Visible = false;
                }
            }
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

        private void Top_Page0_Click(object sender, EventArgs e)
        {
            Paging_now0.Text = (1).ToString();
            Grid_Visible0(int.Parse(Paging_now0.Text));
        }

        private void Previous_Page0_Click(object sender, EventArgs e)
        {
            Paging_now0.Text = (int.Parse(Paging_now0.Text) - 1).ToString();
            Grid_Visible0(int.Parse(Paging_now0.Text));
        }

        private void After_Page0_Click(object sender, EventArgs e)
        {
            Paging_now0.Text = (int.Parse(Paging_now0.Text) + 1).ToString();
            Grid_Visible0(int.Parse(Paging_now0.Text));
        }

        private void End_Page0_Click(object sender, EventArgs e)
        {
            Paging_now0.Text = (int.Parse(Paging_all0.Text)).ToString();
            Grid_Visible0(int.Parse(Paging_now0.Text));
        }

        private void Grid_Visible0(int page)
        {
            for (int i = 1; i < c1FlexGrid0.Rows.Count; i++)
            {
                if ((page - 1) * pagelimit < i && i < page * pagelimit + 1)
                {
                    c1FlexGrid0.Rows[i].Visible = true;
                }
                else
                {
                    c1FlexGrid0.Rows[i].Visible = false;
                }
            }
            set_page_enabled0(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));

        }
        private void set_page_enabled0(int now, int last)
        {
            if (now <= 1)
            {
                Top_Page0.Enabled = false;
                Previous_Page0.Enabled = false;
            }
            else
            {
                Top_Page0.Enabled = true;
                Previous_Page0.Enabled = true;
            }
            if (now >= last)
            {
                End_Page0.Enabled = false;
                After_Page0.Enabled = false;
            }
            else
            {
                End_Page0.Enabled = true;
                After_Page0.Enabled = true;
            }
        }

        private void c1FlexGrid0_AfterSort(object sender, SortColEventArgs e)
        {
            Grid_Visible0(int.Parse(Paging_now0.Text));
        }

        // 工事事務所の選択
        private void c1FlexGrid0_CellChecked(object sender, RowColEventArgs e)
        {
            tableLayoutPanel13.Visible = false;
            for (int i = 1; i < c1FlexGrid0.Rows.Count; i++)
            {
                if (i != e.Row)
                {
                    c1FlexGrid0.SetCellCheck(i, 0, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
                }
                else
                {
                    if (c1FlexGrid0.GetCellCheck(i, 0) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                    {
                        tableLayoutPanel13.Visible = true;
                        string JimushoName = c1FlexGrid0.Rows[i][3].ToString();
                        string JimushoGridNo = i.ToString();

                        c1FlexGrid1.Rows.Count = 1;
                        for (int k = 1; k < Tantoudata.Rows.Count; k++)
                        {
                            //if (Tantoudata.Rows[k][1].ToString() == JimushoName)
                            if (Tantoudata.Rows[k][8].ToString() == JimushoGridNo)
                            {
                                c1FlexGrid1.Rows.Add();
                                c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][2] = Tantoudata.Rows[k][1];
                                c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][3] = Tantoudata.Rows[k][2];
                                c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][4] = Tantoudata.Rows[k][3];
                                c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][5] = Tantoudata.Rows[k][4];
                                c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][6] = Tantoudata.Rows[k][5];
                                c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][7] = Tantoudata.Rows[k][6];
                                c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][8] = Tantoudata.Rows[k][7];
                                c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][9] = Tantoudata.Rows[k][8];
                            }
                        }


                        JimusyoMei.Text = JimushoName;
                        KoujijimushoGridNo.Text = JimushoGridNo;
                        Busho.Text = "";
                        Yakushoku.Text = "";
                        Tantousha.Text = "";
                        Tel.Text = "";
                        Mail.Text = "";
                        Fax.Text = "";
                    }
                }
            }
        }

        private void button_Clear_Click(object sender, EventArgs e)
        {
            Busho.Text = "";
            Yakushoku.Text = "";
            Tantousha.Text = "";
            Tel.Text = "";
            Mail.Text = "";
            Fax.Text = "";
        }

        private void button_Update_Click(object sender, EventArgs e)
        {
            int _row = 0;
            for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
            {
                if (c1FlexGrid1.GetCellCheck(i, 0) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                {
                    _row = i;
                    break;
                }
            }

            set_error("", 0);
            Boolean ErrorFlag = false;
            if (Tantousha.Text == "")
            {
                set_error("担当者を入力して下さい");
                ErrorFlag = true;
            }
            else
            {
                //不具合管理表No1326(1063)
                //担当者名の重複エラーチェックは外す
                //for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
                //{
                //    if (i != _row && c1FlexGrid1.Rows[i][5].ToString() == Tantousha.Text)
                //    {
                //        set_error("担当者が重複しています");
                //        ErrorFlag = true;
                //        break;
                //    }
                //}
            }

            if (ErrorFlag || !check_Mail() || !check_TEL() || !check_FAX())
            {
                return;
            }

            if (_row > 0)
            {
                c1FlexGrid1.Rows[_row][3] = Busho.Text;
                c1FlexGrid1.Rows[_row][4] = Yakushoku.Text;
                c1FlexGrid1.Rows[_row][5] = Tantousha.Text;
                c1FlexGrid1.Rows[_row][6] = Tel.Text;
                c1FlexGrid1.Rows[_row][7] = Fax.Text;
                c1FlexGrid1.Rows[_row][8] = Mail.Text;
                c1FlexGrid1.Rows[_row][9] = KoujijimushoGridNo.Text;
            }
            Change_data();
        }

        private Boolean check_TEL()
        {
            // 0始まり4桁-4桁-4桁 or 11 or 12桁 or 国際番号対応 or 空白
            // No1590対応
            //if (Tel.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(Tel.Text, @"^((0\d{1,4}-\d{1,4}-\d{4})|\+?(\d{10,12})|(\+\d{1,3}-\d{1,2}-\d{1,4}-\d{4})|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            if (Tel.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(Tel.Text, @"^((0\d{1,4}-\d{1,4}-\d{4})|\+?(\d{11,12})|(\+\d{1,3}-\d{1,2}-\d{1,4}-\d{4})|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {
                // E20603:電話番号を正しく入力してください。
                set_error(GlobalMethod.GetMessage("E20603", ""));
                return false;

            }
            return true;
        }

        private Boolean check_FAX()
        {
            // 0始まり4桁-4桁-4桁 or 11 or 12桁 or 空白
            if (Fax.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(Fax.Text, @"^((0\d{1,4}-\d{1,4}-\d{4})|\+?(\d{11,12})|(\+\d{1,3}-\d{1,2}-\d{1,4}-\d{4})|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {
                // E20604:FAX番号を正しく入力してください。
                set_error(GlobalMethod.GetMessage("E20604", ""));
                return false;

            }
            return true;
        }

        private Boolean check_Mail()
        {

            if (Mail.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(Mail.Text, @"^((\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {
                // E10605:メールアドレスを正しく入力してください。
                set_error(GlobalMethod.GetMessage("E10605", ""));
                return false;

            }
            return true;
        }

        // 追加ボタン
        private void button_Insert_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            Boolean ErrorFlag = false;
            if (Tantousha.Text == "")
            {
                set_error("担当者を入力して下さい");
                ErrorFlag = true;
            }
            else
            {
                //不具合管理表No1326(1063)
                //担当者名の重複エラーチェックは外す
                //for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
                //{
                //    if (c1FlexGrid1.Rows[i][5].ToString() == Tantousha.Text)
                //    {
                //        set_error("担当者が重複しています");
                //        ErrorFlag = true;
                //        break;
                //    }
                //}
            }

            if (ErrorFlag || !check_Mail() || !check_TEL() || !check_FAX())
            {
                return;
            }

            c1FlexGrid1.Rows.Add();
            int _row = c1FlexGrid1.Rows.Count - 1;
            c1FlexGrid1.Rows[_row][2] = JimusyoMei.Text;
            c1FlexGrid1.Rows[_row][3] = Busho.Text;
            c1FlexGrid1.Rows[_row][4] = Yakushoku.Text;
            c1FlexGrid1.Rows[_row][5] = Tantousha.Text;
            c1FlexGrid1.Rows[_row][6] = Tel.Text;
            c1FlexGrid1.Rows[_row][7] = Fax.Text;
            c1FlexGrid1.Rows[_row][8] = Mail.Text;
            c1FlexGrid1.Rows[_row][9] = KoujijimushoGridNo.Text;
            Change_data();

            for (int i = 1; i < _row; i++)
            {
                c1FlexGrid1.SetCellCheck(i, 0, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
            }
            c1FlexGrid1.SetCellCheck(_row, 0, CheckEnum.Checked);

            Paging_all.Text = (Math.Ceiling(((double)c1FlexGrid1.Rows.Count - 1) / pagelimit)).ToString();
            Paging_now.Text = Paging_all.Text;
            Grid_Visible(int.Parse(Paging_now.Text));
        }

        private void c1FlexGrid1_BeforeMouseDown(object sender, BeforeMouseDownEventArgs e)
        {
            var hti = ((C1FlexGrid)sender).HitTest(new System.Drawing.Point(e.X, e.Y));
            if (hti.Row > 0 && hti.Column == 1)
            {
                if (GlobalMethod.outputMessage("I10002", "", 1) == DialogResult.OK)
                {
                    c1FlexGrid1.Rows.Remove(hti.Row);
                }
            }
            Change_data();
        }

        private void c1FlexGrid1_AfterSort(object sender, SortColEventArgs e)
        {
            Grid_Visible(int.Parse(Paging_now.Text));
        }

        private void c1FlexGrid1_CellChecked(object sender, RowColEventArgs e)
        {
            for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
            {
                if (i != e.Row)
                {
                    c1FlexGrid1.SetCellCheck(i, 0, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
                }
                else
                {
                    if (c1FlexGrid1.GetCellCheck(i, 0) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                    {
                        Busho.Text = c1FlexGrid1.Rows[i][3].ToString();
                        Yakushoku.Text = c1FlexGrid1.Rows[i][4].ToString();
                        Tantousha.Text = c1FlexGrid1.Rows[i][5].ToString();
                        Tel.Text = c1FlexGrid1.Rows[i][6].ToString();
                        Mail.Text = c1FlexGrid1.Rows[i][8].ToString();
                        Fax.Text = c1FlexGrid1.Rows[i][7].ToString();
                    }
                }
            }
        }

        private void button_Return_Click(object sender, EventArgs e)
        {
            data = Tantoudata;
            this.Close();
        }

        private void Change_data()
        {
            string JimushoName = JimusyoMei.Text;
            string JimushoGridNo = KoujijimushoGridNo.Text;

            int row = 0;

            for (int i = 1; i < Tantoudata.Rows.Count; i++)
            {
                Boolean delflag = true;
                while (delflag)
                {
                    //if (i < Tantoudata.Rows.Count && Tantoudata.Rows[i][1].ToString() == JimushoName)
                    if (i < Tantoudata.Rows.Count && Tantoudata.Rows[i][8].ToString() == JimushoGridNo)
                    {
                        Tantoudata.Rows.Remove(i);

                        //// 削除されたところのindexを取得
                        //if(row == 0)
                        //{
                        //    row = i;
                        //}
                    }
                    else
                    {
                        delflag = false;
                    }
                }
            }
            for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
            {
                Tantoudata.Rows.Add();
                Tantoudata.Rows[Tantoudata.Rows.Count - 1][1] = c1FlexGrid1.Rows[i][2];
                Tantoudata.Rows[Tantoudata.Rows.Count - 1][2] = c1FlexGrid1.Rows[i][3];
                Tantoudata.Rows[Tantoudata.Rows.Count - 1][3] = c1FlexGrid1.Rows[i][4];
                Tantoudata.Rows[Tantoudata.Rows.Count - 1][4] = c1FlexGrid1.Rows[i][5];
                Tantoudata.Rows[Tantoudata.Rows.Count - 1][5] = c1FlexGrid1.Rows[i][6];
                Tantoudata.Rows[Tantoudata.Rows.Count - 1][6] = c1FlexGrid1.Rows[i][7];
                Tantoudata.Rows[Tantoudata.Rows.Count - 1][7] = c1FlexGrid1.Rows[i][8];
                Tantoudata.Rows[Tantoudata.Rows.Count - 1][8] = c1FlexGrid1.Rows[i][9];

                //Tantoudata.Rows.Insert(row + i);
                //Tantoudata.Rows[row + i - 1][1] = c1FlexGrid1.Rows[i][2];
                //Tantoudata.Rows[row + i - 1][2] = c1FlexGrid1.Rows[i][3];
                //Tantoudata.Rows[row + i - 1][3] = c1FlexGrid1.Rows[i][4];
                //Tantoudata.Rows[row + i - 1][4] = c1FlexGrid1.Rows[i][5];
                //Tantoudata.Rows[row + i - 1][5] = c1FlexGrid1.Rows[i][6];
                //Tantoudata.Rows[row + i - 1][6] = c1FlexGrid1.Rows[i][7];
                //Tantoudata.Rows[row + i - 1][7] = c1FlexGrid1.Rows[i][8];
                //Tantoudata.Rows[row + i - 1][8] = c1FlexGrid1.Rows[i][9];
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
    }
}
