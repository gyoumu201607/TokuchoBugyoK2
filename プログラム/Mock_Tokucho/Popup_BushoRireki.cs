﻿using System;
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

namespace TokuchoBugyoK2
{
    public partial class Popup_BushoRireki : Form
    {
        public string[] ReturnValue = new string[10];
        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();
        public string KoujiJimushoMei = "";


        public Popup_BushoRireki()
        {
            InitializeComponent();
        }

        private void Popup_BushoRireki_Load(object sender, EventArgs e)
        {
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            if (KoujiJimushoMei != "")
            {
                item_koujijimushomei.Text = KoujiJimushoMei;
            }
            get_data();
        }

        private void get_data()
        {
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT distinct " +
                        " TanpinHachuubusho " + // 部署
                        ",TanpinTel " +         // 電話
                        ",TanpinFax " +         // FAX
                        //"MadoguchiHachuuKikanmei " +
                        //",TanpinHachuuTantousha " +
                        "FROM TanpinNyuuryoku tn " +
                        "LEFT JOIN MadoguchiJouhou mj on mj.MadoguchiID = tn.MadoguchiID " +
                        "where TanpinHachuubusho != '' AND TanpinTel != '' AND TanpinFax != ''";

                if (item_koujijimushomei.Text != "")
                {
                    cmd.CommandText += "AND MadoguchiHachuuKikanmei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_koujijimushomei.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (item_busho.Text != "")
                {
                    cmd.CommandText += "AND TanpinHachuubusho COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_busho.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (item_tel.Text != "")
                {
                    cmd.CommandText += "AND TanpinTel COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_tel.Text, 1) + "%' ESCAPE '\\' ";
                }
                if (item_fax.Text != "")
                {
                    cmd.CommandText += "AND TanpinFax COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_fax.Text, 1) + "%' ESCAPE '\\' ";
                }

                var sda = new SqlDataAdapter(cmd);
                ListData.Clear();
                sda.Fill(ListData);
            }
            Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / 20)).ToString();
            Paging_now.Text = (1).ToString();
            set_data(1);
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

        private void textChanged(object sender, EventArgs e)
        {
            get_data();
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

        // 閉じる
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid1.HitTest(new Point(e.X, e.Y));

            if (hti.Column == 0 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                ReturnValue[0] = c1FlexGrid1.Rows[_row][_col + 1].ToString(); // 部署
                ReturnValue[1] = c1FlexGrid1.Rows[_row][_col + 2].ToString(); // 電話
                ReturnValue[2] = c1FlexGrid1.Rows[_row][_col + 3].ToString(); // FAX

                //プロンプトを閉じて画面へ戻る
                this.Close();
            }
        }
    }
}
