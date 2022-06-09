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
    public partial class Popup_KoujiJimusyo1 : Form
    {
        GlobalMethod GlobalMethod = new GlobalMethod();
        public C1FlexGrid data = new C1FlexGrid();
        int pagelimit = 20;
        public string AnkenID = "";
        public string JutakuBangou = "";
        public Popup_KoujiJimusyo1()
        {
            InitializeComponent();
        }

        private void Popup_KoujiJimusyo1_Load(object sender, EventArgs e)
        {
            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

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
                HachushaMei.Text = dt.Rows[0][0].ToString();
                HachushaKamei.Text = dt.Rows[0][1].ToString();
            }

            //工事事務所GRIDの初期化
            c1FlexGrid1.Rows.Count = 1;

            //単価契約画面から受け取ったデータの表示
            for (int i = 1; i < data.Rows.Count; i++)
            {
                if (data.Rows[i][2] != null && data.Rows[i][2].ToString() != "")
                {
                    c1FlexGrid1.Rows.Add();
                    for (int k = 2; k < c1FlexGrid1.Cols.Count; k++)
                    {
                        c1FlexGrid1.Rows[i][k] = data.Rows[i][k - 1];
                    }
                }
            }

            data = null;

            Paging_all.Text = (Math.Ceiling(((double)c1FlexGrid1.Rows.Count - 1) / pagelimit)).ToString();
            Grid_Visible(int.Parse(Paging_now.Text));

        }


        private void c1FlexGrid1_CellChecked(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
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
                        JimusyoMei.Text = c1FlexGrid1.Rows[i][3].ToString();
                        JimusyoMeiYomi.Text = c1FlexGrid1.Rows[i][4].ToString();
                        JimusyoBangou.Text = c1FlexGrid1.Rows[i][5].ToString();
                        Yakusyoku.Text = c1FlexGrid1.Rows[i][6].ToString();
                    }
                }
            }
        }

        private void button_insert_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            Boolean ErrorFlag = false;
            if (JimusyoMei.Text == "")
            {
                set_error(GlobalMethod.GetMessage("E20806",""));
                ErrorFlag = true;
            }
            else
            {
                //不具合管理表No1326(1063)
                //工事事務所名の重複エラーチェックは外す
                //for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
                //{
                //    if (c1FlexGrid1.Rows[i][3].ToString() == JimusyoMei.Text)
                //    {
                //        set_error("工事事務所名が重複しています");
                //        ErrorFlag = true;
                //        break;
                //    }
                //}
            }

            if (ErrorFlag)
            {
                return;
            }

            c1FlexGrid1.Rows.Add();
            int _row = c1FlexGrid1.Rows.Count - 1;
            c1FlexGrid1.Rows[_row][3] = JimusyoMei.Text;
            c1FlexGrid1.Rows[_row][4] = JimusyoMeiYomi.Text;
            c1FlexGrid1.Rows[_row][5] = JimusyoBangou.Text;
            c1FlexGrid1.Rows[_row][6] = Yakusyoku.Text;

            for (int i = 1; i < _row; i++)
            {
                c1FlexGrid1.SetCellCheck(i, 0, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
            }
            c1FlexGrid1.SetCellCheck(_row, 0, CheckEnum.Checked);

            Paging_all.Text = (Math.Ceiling(((double)c1FlexGrid1.Rows.Count - 1) / pagelimit)).ToString();
            Paging_now.Text = Paging_all.Text;
            Grid_Visible(int.Parse(Paging_now.Text));
        }

        private void button_update_Click(object sender, EventArgs e)
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
            if (JimusyoMei.Text == "")
            {
                set_error(GlobalMethod.GetMessage("E20806", ""));
                ErrorFlag = true;
            }
            else
            {
                //不具合管理表No1326(1063)
                //工事事務所名の重複エラーチェックは外す
                //for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
                //{
                //    if (i != _row && c1FlexGrid1.Rows[i][3].ToString() == JimusyoMei.Text)
                //    {
                //        set_error("工事事務所名が重複しています");
                //        ErrorFlag = true;
                //        break;
                //    }
                //}
            }

            if (ErrorFlag)
            {
                return;
            }

            if (_row > 0)
            {
                c1FlexGrid1.Rows[_row][3] = JimusyoMei.Text;
                c1FlexGrid1.Rows[_row][4] = JimusyoMeiYomi.Text;
                c1FlexGrid1.Rows[_row][5] = JimusyoBangou.Text;
                c1FlexGrid1.Rows[_row][6] = Yakusyoku.Text;
            }
        }

        private void button_clear_Click(object sender, EventArgs e)
        {
            JimusyoMei.Text = "";
            JimusyoMeiYomi.Text = "";
            JimusyoBangou.Text = "";
            Yakusyoku.Text = "";
        }

        private void button_Return_Click(object sender, EventArgs e)
        {
            this.data = new C1FlexGrid();
            this.data = c1FlexGrid1;
            this.Close();
        }

        private void c1FlexGrid1_BeforeMouseDown(object sender, BeforeMouseDownEventArgs e)
        {
            var hti = ((C1FlexGrid)sender).HitTest(new System.Drawing.Point(e.X, e.Y));
            if (hti.Row > 0 && hti.Column == 1)
            {
                if (GlobalMethod.outputMessage("I10002", "",1) == DialogResult.OK)
                {
                    c1FlexGrid1.Rows.Remove(hti.Row);
                }
            }
        }

        private void item1_TextChanged(object sender, EventArgs e)
        {
            ImeLanguage ImeLanguage = new ImeLanguage();
            if (JimusyoMei.Text != "")
            {
                JimusyoMeiYomi.Text = ImeLanguage.GetYomi(JimusyoMei.Text);
            }
            else
            {
                JimusyoMeiYomi.Text = "";
            }
            ImeLanguage.Dispose();
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

        private void c1FlexGrid1_AfterSort(object sender, SortColEventArgs e)
        {
            Grid_Visible(int.Parse(Paging_now.Text));
        }


        private void yomiTextBox1_CompositionCompleted(object sender, CompositionCompletedEventArgs e)
        {
            JimusyoMeiYomi.Text += e.HanKana;
            if (JimusyoMeiYomi.Text.Length > 150)
            {
                JimusyoMeiYomi.Text = JimusyoMeiYomi.Text.Substring(0,150);
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
