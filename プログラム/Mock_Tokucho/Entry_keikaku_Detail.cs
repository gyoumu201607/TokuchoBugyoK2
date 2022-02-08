using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Deployment.Application;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TokuchoBugyoK2
{
    public partial class Entry_keikaku_Detail : Form
    {
        public string[] UserInfos;
        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();
        public string KeikakuID;

        public Entry_keikaku_Detail()
        {
            InitializeComponent();
        }

        private void Entry_keikaku_Detail_Load(object sender, EventArgs e)
        {
            //ヘッダーユーザ
            label3.Text = UserInfos[3] + "：" + UserInfos[1];
            label20.Text = GlobalMethod.GetCommonValue1("APL_VERSION");
            if (GlobalMethod.GetCommonValue1("BOOT_MODE") == "1")
            {
                label19.Text = GlobalMethod.GetCommonValue2("BOOT_MODE");
            }
            set_combo();

            //表示データ取得
            //仮データを入力しているため、下記内容は修正予定
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            var dt2 = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = @"SELECT 
                    KeikakuUriageNendo,KeikakuTourokubi,KeikakuGyoumuKubun
                    ,KeikakuBangou,KeikakuHachuushaMeiKaMei,KeikakuAnkenMei
                    ,KeikakuZenkaiKeikakuBangou,KeikakuZenkaiAnkenBangou,KeikakuZenkaiJutakuBangou
                    ,KeikakuAnkensu
                    ,KeikakuRakusatsumikomigaku,KeikakuRakusatsumikomigakuJF,KeikakuRakusatsumikomigakuJ,KeikakuRakusatsumikomigakuK
                    ,KeikakuHenkoumikomigaku,KeikakuHenkoumikomigakuJF,KeikakuHenkoumikomigakuJ,KeikakuHenkoumikomigakuK
                    ,KeikakuMikomigaku,KeikakuMikomigakuJF,KeikakuMikomigakuJ,KeikakuMikomigakuK,KeikakuMikomigakuGoukei
                    ,KeikakuShizaiChousa,KeikakuEizen,KeikakuKikiruiChousa,KeikakuKoujiChousahi
                    ,KeikakuSanpaiChousa,KeikakuHokakeChousa,KeikakuShokeihiChousa,KeikakuGenkaBunseki
                    ,KeikakuKijunsakusei,KeikakuKoukyouRoumuhi,KeikakuRoumuhiKoukyouigai,KeikakuSonotaChousabu,KeikakuHaibunGoukei

                    FROM KeikakuJouhou
                    WHERE KeikakuID = " + KeikakuID;

                var sda = new SqlDataAdapter(cmd);

                ListData.Clear();
                sda.Fill(ListData);
            }
            item1_1.Text = ListData.Rows[0][0].ToString() + "年度";
            item1_2.Text = ((DateTime)ListData.Rows[0][1]).ToString("yyyy/MM/dd");
            item1_3.SelectedValue = ListData.Rows[0][2].ToString();
            item1_4.Text = ListData.Rows[0][3].ToString();
            item1_5.Text = ListData.Rows[0][4].ToString();
            item1_6.Text = ListData.Rows[0][5].ToString();
            item1_7.Text = ListData.Rows[0][6].ToString();
            item1_8.Text = ListData.Rows[0][7].ToString();
            item1_9.Text = ListData.Rows[0][8].ToString();
            item1_10.Text = ListData.Rows[0][9].ToString();

            item2_1.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][10]));
            item2_2.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][11]));
            item2_3.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][12]));
            item2_4.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][13]));
            item2_5.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][10] + (decimal)ListData.Rows[0][11] + (decimal)ListData.Rows[0][12] + (decimal)ListData.Rows[0][13]));
            item2_6.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][14]));
            item2_7.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][15]));
            item2_8.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][16]));
            item2_9.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][17]));
            item2_10.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][14] + (decimal)ListData.Rows[0][15] + (decimal)ListData.Rows[0][16] + (decimal)ListData.Rows[0][17]));
            item2_11.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][18]));
            item2_12.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][19]));
            item2_13.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][20]));
            item2_14.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][21]));
            item2_15.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][22]));

            item3_1.Text = ((decimal)ListData.Rows[0][23]).ToString("F2") + "%";
            item3_2.Text = ((decimal)ListData.Rows[0][24]).ToString("F2") + "%";
            item3_3.Text = ((decimal)ListData.Rows[0][25]).ToString("F2") + "%";
            item3_4.Text = ((decimal)ListData.Rows[0][26]).ToString("F2") + "%";
            item3_5.Text = ((decimal)ListData.Rows[0][27]).ToString("F2") + "%";
            item3_6.Text = ((decimal)ListData.Rows[0][28]).ToString("F2") + "%";
            item3_7.Text = ((decimal)ListData.Rows[0][29]).ToString("F2") + "%";
            item3_8.Text = ((decimal)ListData.Rows[0][30]).ToString("F2") + "%";
            item3_9.Text = ((decimal)ListData.Rows[0][31]).ToString("F2") + "%";
            item3_10.Text = ((decimal)ListData.Rows[0][32]).ToString("F2") + "%";
            item3_11.Text = ((decimal)ListData.Rows[0][33]).ToString("F2") + "%";
            item3_12.Text = ((decimal)ListData.Rows[0][34]).ToString("F2") + "%";
            item3_13.Text = ((decimal)ListData.Rows[0][35]).ToString("F2") + "%";

            if (item1_10.Text == "0")
            {
                button3.Enabled = false;
                button3.BackColor = Color.DarkGray;
            }

            if (GlobalMethod.GetIntroductionPhase() <= 1)
            {
                button1.Enabled = false;
                button1.BackColor = Color.DarkGray;
                button5.Visible = false;
            }


            this.Owner.Hide();
        }


        //コンボボックスの内容を設定
        private void set_combo()
        {
            GlobalMethod GlobalMethod = new GlobalMethod();
            //契約区分
            string discript = "GyoumuKubunHyouji";
            string value = "GyoumuNarabijunCD";
            string table = "Mst_GyoumuKubun";
            string where = "GyoumuNarabijunCD < 100 ";
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            DataRow dr = combodt.NewRow();
            combodt.Rows.InsertAt(dr, 0);
            item1_3.DataSource = combodt;
            item1_3.DisplayMember = "Discript";
            item1_3.ValueMember = "Value";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Owner.Show();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Entry_Input form = new Entry_Input();
            form.KeikakuID = KeikakuID;
            form.mode = "keikaku";
            form.UserInfos = this.UserInfos;
            this.Hide();
            form.Show(this);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Entry_keikaku_Search form = new Entry_keikaku_Search();
            form.UserInfos = UserInfos;
            form.Show();
            this.Close();
        }

        private void Entry_keikaku_Detail_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Owner.Visible == false)
            {
                this.Owner.Show();
                this.Owner.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Popup_Anken form = new Popup_Anken();
            form.mode = "keikaku";
            form.keikakubangou = item1_4.Text;

            form.ShowDialog();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Entry_Search form = new Entry_Search();
            form.UserInfos = UserInfos;
            form.Show();
            this.Close();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {

        }
    }
}
