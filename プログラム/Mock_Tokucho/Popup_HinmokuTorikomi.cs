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
    public partial class Popup_HinmokuTorikomi : Form
    {
        public string[] ReturnValue = new string[10];
        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();

        // autoCompListの1,2どちらを使っているか 1:autoCompList1 2:autoCompList2
        int compListFlg = 1;

        AutoCompleteStringCollection autoCompList1;
        AutoCompleteStringCollection autoCompList2;


        public Popup_HinmokuTorikomi()
        {
            InitializeComponent();
        }

        private void Popup_HinmokuTorikomi_Load(object sender, EventArgs e)
        {
            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            autoCompList1 = new AutoCompleteStringCollection();
            autoCompList2 = new AutoCompleteStringCollection();
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
                    "ch.ChousaHinmei " +
                    ",ch.ChousaKikaku " +
                    ",ch.ChousaTanka " +
                    ",ch.ChousaKakaku " +
                    ",ch.MadoguchiID " +
                    "FROM MadoguchiJouhou mj " +
                    "LEFT JOIN ChousaHinmoku ch ON ch.MadoguchiID = mj.MadoguchiID " +
                    "WHERE ISNULL(ch.ChousaDeleteFlag,0) = 0 ";

                if (item_TokuchoBangou.Text != null && item_TokuchoBangou.Text != "")
                {
                    cmd.CommandText += " AND mj.MadoguchiUketsukeBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(item_TokuchoBangou.Text, 1, 0) + "' ";
                }
                else
                {
                    // 特調番号が空の場合は、Gridになにも出さない
                    cmd.CommandText += " AND 1 = 2 ";
                }
                if (item_ComboEdaban.Text != null && item_ComboEdaban.Text != "")
                {
                    cmd.CommandText += " AND mj.MadoguchiUketsukeBangouEdaban COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(item_ComboEdaban.Text, 1, 0) + "' ";
                }
                else
                {
                    // 枝番が空の場合は、Gridになにも出さない
                    cmd.CommandText += " AND 1 = 2 ";
                }
                var sda = new SqlDataAdapter(cmd);
                ListData.Clear();
                sda.Fill(ListData);
            }
            Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / 20)).ToString();
            Paging_now.Text = (1).ToString();
            set_data(1);
        }

        private void set_combo()
        {
            GlobalMethod GlobalMethod = new GlobalMethod();
            //コンボボックスの内容を設定
            var combodt1 = new System.Data.DataTable();
            DataRow dr;

            String discript = "MadoguchiUketsukeBangouEdaban";
            String value = "MadoguchiUketsukeBangouEdaban";
            String table = "MadoguchiJouhou";
            String where = "ISNULL(MadoguchiDeleteFlag,0) = 0 AND MadoguchiUketsukeBangou COLLATE Japanese_XJIS_100_CI_AS_SC =  N'" + item_TokuchoBangou.Text + "' ";

            //コンボボックスデータ取得
            combodt1 = GlobalMethod.getData(discript, value, table, where);
            if (combodt1 != null)
            {
                dr = combodt1.NewRow();
                combodt1.Rows.InsertAt(dr, 0);
            }
            item_ComboEdaban.DataSource = combodt1;
            item_ComboEdaban.DisplayMember = "Discript";
            item_ComboEdaban.ValueMember = "Value";
            get_data();
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

        // 特調番号
        private void item_TokuchoBangou_TextChanged(object sender, EventArgs e)
        {
            // AccessViolationException 回避対策
            item_TokuchoBangou.AutoCompleteMode = AutoCompleteMode.None;
            item_TokuchoBangou.AutoCompleteSource = AutoCompleteSource.None;
            item_TokuchoBangou.AutoCompleteCustomSource = null;

            if(item_TokuchoBangou.Text.Length >= 1) 
            { 
                GlobalMethod GlobalMethod = new GlobalMethod();
                //コンボボックスの内容を設定
                DataTable combodt1 = new System.Data.DataTable();
                DataRow dr;

                String discript = "MadoguchiUketsukeBangou";
                String value = "MadoguchiUketsukeBangou";
                String table = "MadoguchiJouhou";
                String where = "ISNULL(MadoguchiDeleteFlag,0) = 0 AND MadoguchiUketsukeBangou COLLATE Japanese_XJIS_100_CI_AS_SC like N'" + item_TokuchoBangou.Text + "%' ESCAPE '\\' ";

                // AutoComplete のバグで、メモリ破損エラーが出る
                // Windows Forms does not protect the AutoCompleteCustomSource object from being replaced while it is being enumerated by a background thread created by autocomplete.
                // Windows Formsでは、オートコンプリートによって作成されたバックグラウンドスレッドで列挙されている間に、AutoCompleteCustomSourceオブジェクトが置き換えられないように保護されていません。

                // AutoComplete を2つ作って置き換える方法でエラーを回避する
                //if (compListFlg == 1)
                //{
                //    compListFlg = 2;
                //    autoCompList2.Clear();
                //}
                //else
                //{
                //    compListFlg = 1;
                //    autoCompList1.Clear();
                //}
                autoCompList1.Clear();

                //コンボボックスデータ取得
                combodt1 = GlobalMethod.getData(discript, value, table, where);
                if (combodt1 != null)
                {
                    dr = combodt1.NewRow();
                    combodt1.Rows.InsertAt(dr, 0);
                    for (int i = 0; i < combodt1.Rows.Count; i++)
                    {
                        //if (compListFlg == 1)
                        //{
                        //    autoCompList1.Add(combodt1.Rows[i][0].ToString());
                        //}
                        //else
                        //{
                        //    autoCompList2.Add(combodt1.Rows[i][0].ToString());
                        //}
                        autoCompList1.Add(combodt1.Rows[i][0].ToString());
                    }
                }

                // ここで設定する為、デザインからは設定を外す
                item_TokuchoBangou.AutoCompleteMode = AutoCompleteMode.Suggest;
                item_TokuchoBangou.AutoCompleteSource = AutoCompleteSource.CustomSource;
                item_TokuchoBangou.AutoCompleteCustomSource = autoCompList1;
                //if (compListFlg == 1)
                //{
                //    item_TokuchoBangou.AutoCompleteCustomSource = autoCompList1;
                //}
                //else
                //{
                //    item_TokuchoBangou.AutoCompleteCustomSource = autoCompList2;
                //}
            }

            set_combo();
        }

        // 枝番
        private void item_ComboEdaban_TextChanged(object sender, EventArgs e)
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

        //　キャンセル
        private void item_BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 登録
        private void item_BtnTouroku_Click(object sender, EventArgs e)
        {
            string MadoguchiID = "";
            // 実行結果に1を追加
            ReturnValue[0] = "1";
            for (int i = 0; i < ListData.Rows.Count; i++)
            {
                MadoguchiID = ListData.Rows[0][4].ToString();
                break;
            }
            ReturnValue[1] = MadoguchiID;
            this.Close();
        }
    }
}
