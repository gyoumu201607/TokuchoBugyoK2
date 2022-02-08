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

namespace TokuchoBugyoK2
{
    public partial class Popup_TokuchoNo : Form
    {
        public string[] ReturnValue = new string[10];
        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();
        public string tokuchouNo = "";

        public Popup_TokuchoNo()
        {
            InitializeComponent();
        }

        private void Popup_Gijutsusya_Load(object sender, EventArgs e)
        {
            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            src_1.Text = tokuchouNo;
            c1FlexGrid1.AutoSizeCol(2, 4);
            get_data();
        }


        private void get_data()
        {
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT " +
                    "mj.MadoguchiJutakuBushoCD " + // 0:部所
                    ",mb.ShibuMei + ' ' + IsNull(mb.KaMei,'') " + // 1:支部名 + 課名
                    ",case when MadoguchiUketsukeBangouEdaban is null OR MadoguchiUketsukeBangouEdaban = '' then MadoguchiUketsukeBangou " + // 2:特調番号 + 枝番（存在すれば）
                    "else MadoguchiUketsukeBangou + '-' + MadoguchiUketsukeBangouEdaban end as tokuchoNo " +
                    //",mj.MadoguchiJutakuBangou " +  // 3:特調番号
                    //",mj.MadoguchiUketsukeBangou " +  // 3:特調番号 画面上は受託番号だが、表示するのは特調番号（MadoguchiUketsukeBangou）
                    //",mj.MadoguchiJutakuBangouEdaban " + // 4:特調枝番
                    //",mj.MadoguchiJutakuBangou " + // 3:受託番号
                    ",CASE mj.MadoguchiJutakuBangouEdaban WHEN ''  THEN mj.MadoguchiJutakuBangou ELSE mj.MadoguchiJutakuBangou + '-' + mj.MadoguchiJutakuBangouEdaban END AS JutakuBangou " + // 3:受託番号
                    ",mj.MadoguchiUketsukeBangouEdaban " + // 4:特調枝番
                    ",mj.MadoguchiKanriBangou " + // 5:管理番号
                    ",mj.MadoguchiTourokubi " + // 6:登録日
                    "FROM MadoguchiJouhou mj INNER JOIN Mst_Busho mb ON  mj.MadoguchiJutakuBushoCD = mb.GyoumuBushoCD " +
                    "WHERE mj.MadoguchiID > 0 AND MadoguchiDeleteFlag != 1 ";

                if (src_1.Text != "")
                {
                    cmd.CommandText += "AND CONCAT(mj.MadoguchiUketsukeBangou,'-', mj.MadoguchiUketsukeBangouEdaban) COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_1.Text, 1) + "%' ESCAPE '\\' ";
                }

                cmd.CommandText += "ORDER BY tokuchoNo ";
                var sda = new SqlDataAdapter(cmd);
                ListData.Clear();
                sda.Fill(ListData);
            }
            Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / 20)).ToString();
            Paging_now.Text = (1).ToString();
            set_data(1);
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
                for (int i = 0; i < c1FlexGrid1.Cols.Count; i++)
                {
                    c1FlexGrid1[r + 1, i] = ListData.Rows[startrow + r][i];
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

        // 特調番号検索
        private void src_1_TextChanged(object sender, EventArgs e)
        {
            get_data();
        }

        private void src_1_KeyDown(object sender, KeyEventArgs e)
        {
            get_data();
        }
    }
}
