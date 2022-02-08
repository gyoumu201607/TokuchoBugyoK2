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

namespace TokuchoBugyoK2
{
    public partial class Popup_Hachusha : Form
    {
        public string[] ReturnValue = new string[10];
        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();

        public Popup_Hachusha()
        {
            InitializeComponent();
        }


        private void Popup_Hachusha_Load(object sender, EventArgs e)
        {
            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

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
                    "HachushaCD " +
                    ",HachushaKubun1Mei " +
                    ",HachushaKubun2Mei " +
                    ",TodouhukenMei " +
                    ",HachushaMei " +
                    "FROM Mst_Hachusha " +
                    "LEFT JOIN Mst_HachushaKubun1 ON Mst_Hachusha.HachushaKubun1CD = Mst_HachushaKubun1.HachushaKubun1CD " +
                    "LEFT JOIN Mst_HachushaKubun2 ON Mst_Hachusha.HachushaKubun2CD = Mst_HachushaKubun2.HachushaKubun2CD " +
                    "LEFT JOIN Mst_Todouhuken ON Mst_Hachusha.TodouhukenCD = Mst_Todouhuken.TodouhukenCD " +
                    "WHERE HachushaHyoujiFlg = 1 ";

                if (src_1.Text != "")
                {
                    cmd.CommandText += " AND HachushaCD COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_1.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (src_2.Text != "")
                {
                    cmd.CommandText += " AND HachushaKubun1Mei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_2.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (src_3.Text != "")
                {
                    cmd.CommandText += " AND HachushaKubun2Mei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_3.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (src_4.Text != "")
                {
                    cmd.CommandText += " AND TodouhukenMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_4.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (src_5.Text != "")
                {
                    cmd.CommandText += " AND HachushaMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_5.Text, 1) + "%' ESCAPE '\\' ";
                }
                var sda = new SqlDataAdapter(cmd);
                ListData.Clear();
                sda.Fill(ListData);
            }
            Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / 20)).ToString();
            Paging_now.Text = (1).ToString();
            set_data(1);
        }

        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {

            var hti = this.c1FlexGrid1.HitTest(new System.Drawing.Point(e.X, e.Y));

            if (hti.Column == 0 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                ReturnValue[0] = c1FlexGrid1.Rows[_row][_col + 1].ToString();
                ReturnValue[1] = c1FlexGrid1.Rows[_row][_col + 2].ToString();
                ReturnValue[2] = c1FlexGrid1.Rows[_row][_col + 3].ToString();
                ReturnValue[3] = c1FlexGrid1.Rows[_row][_col + 4].ToString();
                ReturnValue[4] = c1FlexGrid1.Rows[_row][_col + 5].ToString();
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Top_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (1).ToString();
            set_data(int.Parse(Paging_now.Text));
        }

        private void Previous_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_now.Text) - 1).ToString();
            set_data(int.Parse(Paging_now.Text));
        }

        private void After_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_now.Text) + 1).ToString();
            set_data(int.Parse(Paging_now.Text));
        }

        private void End_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_all.Text)).ToString();
            set_data(int.Parse(Paging_now.Text));
        }
        private void set_data(int pagenum)
        {
            c1FlexGrid1.Rows.Count = 1;
            c1FlexGrid1.AllowAddNew = true;
            int viewnum = 20;
            int startrow = (pagenum - 1) * viewnum;
            int addnum = ListData.Rows.Count - startrow;
            if (addnum > viewnum)
            {
                addnum = viewnum;
            }
            for (int r = 0; r < addnum; r++)
            {
                c1FlexGrid1.Rows.Add();
                for (int i = 0; i < c1FlexGrid1.Cols.Count - 1; i++)
                {
                    if (i < ListData.Columns.Count)
                    {
                        c1FlexGrid1[r + 1, i + 1] = ListData.Rows[startrow + r][i];
                    }
                }

            }
            c1FlexGrid1.AllowAddNew = false;
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

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            get_data();
        }

    }
}
