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
    public partial class Popup_Keikaku : Form
    {
        public string[] ReturnValue = new string[10];
        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();

        public string gyoumuBushoCD = "";
        public string nendo;
        
        public Popup_Keikaku()
        {
            InitializeComponent();
        }

        private void Popup_Keikaku_Load(object sender, EventArgs e)
        {
            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            // ホイール制御
            this.src_1.MouseWheel += item_MouseWheel; // 売上年度

            set_combo();

            // 部所CDが渡ってきていたら
            if(gyoumuBushoCD != "")
            {
                setShibuMei();
            }

            get_data();
        }

        // 計画部所支部、計画課所支部をセットする
        private void setShibuMei()
        {
            GlobalMethod GlobalMethod = new GlobalMethod();
            String discript = "KaMei";
            String value = "ShibuMei ";
            String table = "Mst_Busho";
            String where = "";

            //int FromNendo = 0;
            //int ToNendo = 0;
            //if (int.TryParse(nendo, out FromNendo))
            //{
            //    ToNendo = FromNendo + 1;
            //}
            //else
            //{
            //    FromNendo = DateTime.Today.Year;
            //    ToNendo = FromNendo + 1;
            //}
            //where = "GyoumuBushoCD < '999990' AND BushoNewOld <= 1 AND ISNULL(BushoDeleteFlag,0) = 0 " +
            //        "AND NOT GyoumuBushoCD LIKE '121%' " + 
            //        "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
            //        "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";

            // 渡ってきた部所CDでそのまま検索
            where = "GyoumuBushoCD = '" + gyoumuBushoCD + "' ";

            //コンボボックスデータ取得
            DataTable dt = GlobalMethod.getData(discript, value, table, where);

            // 部所支部、課所支部をセットする
            if (dt != null && dt.Rows.Count > 0)
            {
                src_5.Text = dt.Rows[0][0].ToString();
                item_KeikakuKashoShibu.Text = dt.Rows[0][1].ToString();
            }

        }

        private void set_combo()
        {
            GlobalMethod GlobalMethod = new GlobalMethod();
            //売上年度
            String discript = "NendoSeireki";
            String value = "NendoID ";
            String table = "Mst_Nendo";

            // 売上年度は年度マスタのデータを全て表示
            //String where = "NendoID <= YEAR(GETDATE()) AND NendoID >= YEAR(GETDATE()) - 3";
            String where = "";
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            src_1.DataSource = combodt;
            src_1.DisplayMember = "Discript";
            src_1.ValueMember = "Value";
            /*
            discript = "NendoSeireki";
            value = "NendoID ";
            table = "Mst_Nendo";
            where = "Nendo_Sdate <= GETDATE() AND Nendo_EDate >= GETDATE()";
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


        }

        private void get_data()
        {
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            var dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {

                int FromNendo = 0;
                int ToNendo = 0;
                if (int.TryParse(nendo, out FromNendo))
                {
                    ToNendo = FromNendo + 1;
                }
                else
                {
                    FromNendo = DateTime.Today.Year;
                    ToNendo = FromNendo + 1;
                }

                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT " +
                    //"KeikakuUriageNendo,KeikakuBangou,KeikakuHachuushaMei,KeikakuAnkenMei,KeikakuBushoShibuMei,KeikakuZenkaiAnkenBangou,KeikakuZenkaiJutakuBangou,KeikakuGyoumuKubunMei" +
                    "KeikakuUriageNendo,KeikakuBangou,KeikakuHachuushaMeiKaMei,KeikakuAnkenMei,KeikakuBushoShibuMei,KeikakuZenkaiAnkenBangou,KeikakuZenkaiJutakuBangou,KeikakuGyoumuKubunMei " +
                    //",mb.GyoumuBushoCD" + 
                    //" FROM KeikakuJouhou WHERE KeikakuBangou <> '' ";
                    " FROM KeikakuJouhou kj " +
                    "LEFT JOIN Mst_Busho mb on mb.BushoShibuCD = kj.KeikakuBushoShibuCD " +
                    "AND mb.KashoShibuCD = kj.KeikakuKashoShibuCD " +
                    //"AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                    //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
                    "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' )";


                cmd.CommandText += "WHERE KeikakuBangou <> '' ";

                if (src_1.Text != "")
                {
                    cmd.CommandText += "AND KeikakuUriageNendo COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(src_1.SelectedValue.ToString(), 0) + "' ";
                }
                if (src_2.Text != "")
                {
                    cmd.CommandText += "AND KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_2.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (src_3.Text != "")
                {
                    cmd.CommandText += "AND KeikakuHachuushaMeiKaMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_3.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (src_4.Text != "")
                {
                    cmd.CommandText += "AND KeikakuAnkenMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_4.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (src_5.Text != "")
                {
                    //cmd.CommandText += "AND KeikakuBushoShibuMei LIKE '%" + GlobalMethod.ChangeSqlText(src_5.Text, 1) + "%' ESCAPE '\\' ";
                    cmd.CommandText += "AND mb.ShibuMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_5.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (item_KeikakuKashoShibu.Text != "")
                {
                    //cmd.CommandText += "AND KeikakuKashoShibuMei LIKE '%" + GlobalMethod.ChangeSqlText(item_KeikakuKashoShibu.Text, 1) + "%' ESCAPE '\\' ";
                    cmd.CommandText += "AND mb.KaMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_KeikakuKashoShibu.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (src_6.Text != "")
                {
                    cmd.CommandText += "AND KeikakuZenkaiAnkenBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_6.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (src_7.Text != "")
                {
                    cmd.CommandText += "AND KeikakuZenkaiJutakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_7.Text, 1) + "%' ESCAPE '\\' ";
                }
                cmd.CommandText += "AND KeikakuBangou not like '%-P9%' ";


                cmd.CommandText += "ORDER BY KeikakuBangou DESC";
                var sda = new SqlDataAdapter(cmd);
                ListData.Clear();
                sda.Fill(ListData);

                Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / 20)).ToString();
                Paging_now.Text = (1).ToString();
                set_data(1);
            }
        }

        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {

            var hti = this.c1FlexGrid1.HitTest(new System.Drawing.Point(e.X, e.Y));

            if (hti.Column == 0 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                ReturnValue[0] = c1FlexGrid1.Rows[_row][_col + 2].ToString();
                ReturnValue[1] = c1FlexGrid1.Rows[_row][_col + 4].ToString();

                //ReturnValue[2] = c1FlexGrid1.Rows[_row][_col + 9].ToString(); // 計画課所支部CD（GyoumuBushoCD）
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
                    c1FlexGrid1[r + 1, i + 1] = ListData.Rows[startrow + r][i];
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

        private void src_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void src_1_TextChanged(object sender, EventArgs e)
        {
            get_data();
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
        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        private void src_1_SelectedIndexChanged(object sender, EventArgs e)
        {
            get_data();
        }
    }
}
