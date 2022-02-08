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
    public partial class Popup_TanpinTantousya : Form
    {
        DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();
        public string[] ReturnValue = new string[6];
        public string KoujiJimusyoMei = "";
        //public string AnkenJouhouID = "";
        public string MadoguchiID = "";
        public string TankaKeiyakuID = "0";
        private int pagelimit = 20;

        public Popup_TanpinTantousya()
        {
            InitializeComponent();
        }

        private void Popup_TanpinTantousya_Load(object sender, EventArgs e)
        {
            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            if (KoujiJimusyoMei != "")
            {
                src_1.Text = KoujiJimusyoMei;
            }
            // 単価契約IDを取得する
            DataTable Tanpin_Dt = new DataTable();
            Tanpin_Dt = GlobalMethod.getData("TanpinGyoumuCD", "TanpinGyoumuCD", "TanpinNyuuryoku", "MadoguchiID = " + MadoguchiID);
            if (Tanpin_Dt.Rows.Count > 0 && Tanpin_Dt.Rows[0][0] != null)
            {
                TankaKeiyakuID = Tanpin_Dt.Rows[0][0].ToString();
            }
            get_data();
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
        private void src_TextChanged(object sender, EventArgs e)
        {
            get_data();
        }
        private void get_data()
        {
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            var dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT " +
                  "distinct " + // 1024 担当者が2つ出る クエリとしては間違ってないが、別の単価契約で同じ人を登録したからだと思われる、その為、重複を除外
                  "KoujijimushoMei,KoujiTantoushaBusho,KoujiTantoushaYakushoku,KoujiTantoushaMei,KoujiTantoushaTEL,KoujiTantoushaFAX,KoujiTantoushaMail " +
                  "FROM Mst_KoujijimushoTantousha " +
                  "LEFT JOIN Mst_Koujijimusho ON Mst_KoujijimushoTantousha.KoujijimushoID = Mst_Koujijimusho.KoujijimushoID " +
                  //"LEFT JOIN TanpinNyuuryoku ON Mst_Koujijimusho.TankaKeiyakuID = TanpinNyuuryoku.TanpinGyoumuCD " +
                  //"WHERE TanpinNyuuryoku.MadoguchiID = '" + MadoguchiID + "' ";
                  "WHERE KoujiTantoushaDeleteFlag <> 1 " +
                  " AND TankaKeiyakuID = " + TankaKeiyakuID + " ";

                if (src_1.Text != "")
                {
                    cmd.CommandText += " AND KoujijimushoMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_1.Text, 0) + "%' ESCAPE'\\' ";
                }
                if (src_2.Text != "")
                {
                    cmd.CommandText += " AND KoujiTantoushaBusho COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_2.Text, 1) + "%' ESCAPE'\\' ";
                }
                if (src_3.Text != "")
                {
                    cmd.CommandText += " AND KoujiTantoushaMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_3.Text, 1) + "%' ESCAPE'\\' ";
                }
                if (src_4.Text != "")
                {
                    cmd.CommandText += " AND KoujiTantoushaMail COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_4.Text, 1) + "%' ESCAPE'\\' ";
                }
                cmd.CommandText += " ORDER BY KoujijimushoMei,KoujiTantoushaBusho,KoujiTantoushaMei";
                Console.WriteLine(cmd.CommandText);
                var sda = new SqlDataAdapter(cmd);
                ListData.Clear();
                //if (AnkenJouhouID != "")
                if (MadoguchiID != "")
                {
                    sda.Fill(ListData);
                }
            }
            Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / pagelimit)).ToString();
            Paging_now.Text = (1).ToString();
            set_data();
            Grid_Visible(1);
        }
        private void set_data()
        {
            c1FlexGrid1.Rows.Count = 1;
            for (int r = 0; r < ListData.Rows.Count; r++)
            {
                c1FlexGrid1.Rows.Add();
                for (int i = 0; i < c1FlexGrid1.Cols.Count - 1; i++)
                {
                    if (ListData.Columns.Count > i)
                    {
                        c1FlexGrid1[r + 1, i + 1] = ListData.Rows[r][i];
                    }
                }
            }
            c1FlexGrid1.AllowAddNew = false;
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
        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid1.HitTest(new System.Drawing.Point(e.X, e.Y));

            if (hti.Column == 0 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                ReturnValue[0] = c1FlexGrid1.Rows[_row][2].ToString();     // KoujiTantoushaBusho
                ReturnValue[1] = c1FlexGrid1.Rows[_row][3].ToString();     // KoujiTantoushaYakushoku
                ReturnValue[2] = c1FlexGrid1.Rows[_row][4].ToString();     // KoujiTantoushaMei
                ReturnValue[3] = c1FlexGrid1.Rows[_row][5].ToString();     // KoujiTantoushaTEL
                ReturnValue[4] = c1FlexGrid1.Rows[_row][6].ToString();     // KoujiTantoushaFAX
                ReturnValue[5] = c1FlexGrid1.Rows[_row][7].ToString();     // KoujiTantoushaMail
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
