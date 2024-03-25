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
    public partial class Popup_ChousainList : Form
    {
        public string[] ReturnValue = new string[10];
        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();
        public string nendo;
        private int FromNendo;
        private int ToNendo;
        public string Busho = null;
        public string program;

        public Popup_ChousainList()
        {
            InitializeComponent();
        }

        private void Popup_Gijutsusya_Load(object sender, EventArgs e)
        {

            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            // ホイール制御
            this.src_1.MouseWheel += item_MouseWheel; // 部所名

            if (int.TryParse(nendo, out FromNendo))
            {
                ToNendo = FromNendo + 1;
            }
            else
            {
                FromNendo = DateTime.Today.Year;
                ToNendo = FromNendo + 1;
            }
            set_combo_shibu(FromNendo.ToString());

            if (Busho != null)
            {
                src_1.SelectedValue = Busho;
            }

            get_data();
        }


        private void set_combo_shibu(string nendo)
        {
            //受託課所支部
            //SQL変数
            string discript = "";
            string value = "";
            string table = "";
            string where = "";

            if ("madoguchi".Equals(program))
            {
                discript = "Mst_Busho.BushokanriboKamei ";
                value = "Mst_Busho.GyoumuBushoCD ";
                table = "Mst_Busho";
                where = "GyoumuBushoCD < '999990' AND BushoNewOld <= 1 AND ISNULL(BushoDeleteFlag,0) = 0 " +
                        //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' ";
                        "AND BushokanriboKamei != '' ";
            }
            else
            {
                discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
                value = "Mst_Busho.GyoumuBushoCD ";
                table = "Mst_Busho";
                where = "GyoumuBushoCD < '999990' AND BushoNewOld <= 1 AND ISNULL(BushoDeleteFlag,0) = 0 " +
                        //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' ";
                        "AND NOT GyoumuBushoCD LIKE '121%' ";
            }

            //窓口ミハルからの呼び出しの場合
            if ("madoguchi".Equals(program))
            {
                where += "AND BushoMadoguchiHyoujiFlg = 1 ";
            }
            else
            {
                where += "AND BushoEntryHyoujiFlg = 1 ";
            }

            int FromNendo;
            if (int.TryParse(nendo, out FromNendo))
            {
                int ToNendo = int.Parse(nendo) + 1;
                //where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
                where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' )";
            }

            //窓口ミハルからの呼び出しの場合
            if ("madoguchi".Equals(program))
            {
                where += "ORDER BY BushoMadoguchiNarabijun ";
            }

            Console.WriteLine(where);
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            //if ("madoguchi".Equals(program))
            //{ 
            //    // 窓口ミハルの場合、空行があるので、追加（エントリくん要確認）
            //    DataRow dr;
            //    if (combodt != null)
            //    {
            //        dr = combodt.NewRow();
            //        combodt.Rows.InsertAt(dr, 0);
            //    }
            //}

            // 713対応で部所は空を許容する
            DataRow dr;
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }

            src_1.DataSource = combodt;
            src_1.DisplayMember = "Discript";
            src_1.ValueMember = "Value";
        }


        private void get_data()
        {
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT " +
                    "KojinCD " +
                    ",ChousainMei " +
                    ",Mst_Chousain.GyoumuBushoCD " +
                    ",ShibuMei " +
                    ",ChousaShozoku " +
                    ",ShozokuRyaku " +
                    ",BushoShibuCD " + //支部コード
                    ",KashoShibuCD " + //課コード
                    "FROM Mst_Chousain LEFT JOIN Mst_Busho ON Mst_Chousain.GyoumuBushoCD = Mst_Busho.GyoumuBushoCD " +
                    "WHERE RetireFLG <> 1 AND TokuchoFLG = 1 ";

                //今日日付の年度データを検索する際は、今日有効な調査員を表示
                if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(FromNendo, 4, 1))
                {
                    cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                    "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today + "' ) ";
                }
                else
                {
                    //cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                    //"AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                    cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                    "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + FromNendo + "/4/1' ) ";
                }

                if (src_1.Text != "")
                {
                    // 工期開始年度が2021年度未満の場合、旧積シス（127910）を選択していたら、情報システム部 積算情報課（128416）を表示する
                    if (FromNendo < 2021)
                    {
                        if (src_1.SelectedValue != null && "127910".Equals(src_1.SelectedValue.ToString()))
                        {
                            cmd.CommandText += "AND Mst_Chousain.GyoumuBushoCD = '128416' ";
                        }
                        else
                        {
                            //No1702 業務部署コードの下1桁ゼロ切り捨ての前方一致を完全一致に変更
                            //cmd.CommandText += "AND Mst_Chousain.GyoumuBushoCD COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'" + GlobalMethod.ChangeSqlText(src_1.SelectedValue.ToString().TrimEnd('0'), 1) + "%' ESCAPE '\\' ";
                            cmd.CommandText += "AND Mst_Chousain.GyoumuBushoCD COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'" + GlobalMethod.ChangeSqlText(src_1.SelectedValue.ToString(), 1) + "' ESCAPE '\\' ";
                        }
                    }
                    else
                    {
                        //No1702 業務部署コードの下1桁ゼロ切り捨ての前方一致を完全一致に変更
                        //cmd.CommandText += "AND Mst_Chousain.GyoumuBushoCD COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'" + GlobalMethod.ChangeSqlText(src_1.SelectedValue.ToString().TrimEnd('0'), 1) + "%' ESCAPE '\\' ";
                        cmd.CommandText += "AND Mst_Chousain.GyoumuBushoCD COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'" + GlobalMethod.ChangeSqlText(src_1.SelectedValue.ToString(), 1) + "' ESCAPE '\\' ";
                    }
                }
                if (src_2.Text != "")
                {
                    cmd.CommandText += "AND Mst_Chousain.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_2.Text, 1) + "%' ESCAPE '\\' ";
                }

                //cmd.CommandText += "ORDER BY ChousainMei ";
                // 415：集計表画面の部所の表示で、部所の並び順が入っていない為、本部などで検索するとバラバラに表示される。 対応
                //cmd.CommandText += "ORDER BY Mst_Busho.GyoumuBushoCD,Mst_Chousain.KojinCD ";
                if ("madoguchi".Equals(program))
                {
                    // 窓口ミハルは部所名を空で検索できるので、以下の条件を追加し、窓口データに絞り込む
                    cmd.CommandText += "AND BushoMadoguchiHyoujiFlg = 1 AND BushokanriboKamei != '' ";
                    cmd.CommandText += "ORDER BY Mst_Busho.BushoMadoguchiNarabijun,Mst_Chousain.KojinCD ";
                }
                else
                {
                    cmd.CommandText += "ORDER BY Mst_Chousain.KojinCD ";
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
            var hti = this.c1FlexGrid1.HitTest(new Point(e.X, e.Y));

            //if (hti.Column == 0 & hti.Row != 0)
            // グレーゾーンを押下時には、hti.Row は -1がくる
            if (hti.Column == 0 & hti.Row > 0)
            {
                //選択したデータを配列に格納
                var _row = hti.Row;
                var _col = hti.Column;
                ReturnValue[0] = c1FlexGrid1.Rows[_row][_col + 1].ToString();//調査員CD
                ReturnValue[1] = c1FlexGrid1.Rows[_row][_col + 2].ToString();//調査員名

                // 工期開始年度が2021年度未満の場合、情報システム部 積算情報課（128416）が選択されていた場合、旧シス（127910）を返す
                if (FromNendo < 2021) 
                {
                    if ("128416".Equals(c1FlexGrid1.Rows[_row][_col + 3].ToString())) 
                    { 
                        ReturnValue[2] = "127910";//部所コード
                    }
                    else
                    {
                        ReturnValue[2] = c1FlexGrid1.Rows[_row][_col + 3].ToString();//部所コード
                    }
                }
                else
                {
                    ReturnValue[2] = c1FlexGrid1.Rows[_row][_col + 3].ToString();//部所コード
                }

                ReturnValue[3] = c1FlexGrid1.Rows[_row][_col + 7].ToString();//支部コード
                ReturnValue[4] = c1FlexGrid1.Rows[_row][_col + 8].ToString();//課コード
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

        //private void src_1_TextChanged(object sender, EventArgs e)
        //{
        //    get_data();
        //}

        private void src_2_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    get_data();
            //}
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

        // コンボボックスの選択後のイベントTextChangedで拾う
        private void src_1_TextChanged(object sender, EventArgs e)
        {
            get_data();
        }

        // 職員名
        private void src_2_TextChanged(object sender, EventArgs e)
        {
            get_data();
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
