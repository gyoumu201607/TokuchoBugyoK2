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
using System.Drawing.Imaging;
using System.Deployment.Application;
using Microsoft.VisualBasic.ApplicationServices;

namespace TokuchoBugyoK2
{
    public partial class TopMenu : Form
    {
        public string[] UserInfos;
        DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();


        public TopMenu()
        {
            InitializeComponent();

        }


        private void TopMenu_Load(object sender, EventArgs e)
        {

            //レイアウトロジックを停止する
            this.SuspendLayout();
            label1.Text = UserInfos[3] + "：" + UserInfos[1];
            label6.Text = GlobalMethod.GetCommonValue1("APL_VERSION");
            if (GlobalMethod.GetCommonValue1("BOOT_MODE") == "1")
            {
                label7.Text = GlobalMethod.GetCommonValue2("BOOT_MODE");
            }

            //DB接続情報の取得
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            //お知らせ情報の取得
            Get_Date();


            int PhaseFLG = GlobalMethod.GetIntroductionPhase();
            if (PhaseFLG <= 1)
            {
                button4.Visible = false;
                button3.Visible = false;
                button1.Visible = false;
            }
            else if(PhaseFLG <= 2)
            {
                button4.Visible = false;
                button3.Visible = false;
                button1.Visible = false;
            }

            if (int.Parse(UserInfos[4]) < 1)
            {
                button2.Enabled = false;
                button2.BackColor = Color.DarkGray;
            }

            // フラグにより業務日報ボタンの活性、非活性を切り替える
            string nippouBtnFlg = "0";

            // 0:非活性、1：活性
            nippouBtnFlg = GlobalMethod.GetCommonValue1("NIPPOU_BUTTON_FLGA");

            // マスタにフラグが存在しない場合は、非活性
            if(nippouBtnFlg == null)
            {
                nippouBtnFlg = "0";
            }
            if(nippouBtnFlg == "1")
            {
                // 活性
                button3.Enabled = true;
            }
            else
            {
                // 非活性
                button3.Enabled = false;
                button3.BackColor = Color.DarkGray;
            }

            // フラグにより特調野郎ボタンの活性、非活性を切り替える
            string tokuchoBtnFlg = "0";

            // 0:非活性、1：活性
            tokuchoBtnFlg = GlobalMethod.GetCommonValue1("TOKUCHO_BUTTON_FLGA");

            // マスタにフラグが存在しない場合は、非活性
            if (tokuchoBtnFlg == null)
            {
                tokuchoBtnFlg = "0";
            }
            if (tokuchoBtnFlg == "1")
            {
                // 活性
                button4.Enabled = true;
            }
            else
            {
                // 非活性
                button4.Enabled = false;
                button4.BackColor = Color.DarkGray;
            }

            // フラグにより単価契約ボタンの活性、非活性を切り替える
            string tankaBtnFlg = "0";

            // 0:非活性、1：活性
            tankaBtnFlg = GlobalMethod.GetCommonValue1("TANKA_BUTTON_FLGA");

            // マスタにフラグが存在しない場合は、非活性
            if (tankaBtnFlg == null)
            {
                tankaBtnFlg = "0";
            }
            if (tankaBtnFlg == "1")
            {
                // 活性
                button1.Enabled = true;
            }
            else
            {
                // 非活性
                button1.Enabled = false;
                button1.BackColor = Color.DarkGray;
            }

            this.Owner.Hide();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void Get_Date()
        {
            try
            {
                //DB接続情報の取得
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT  " +
                        "T_INFORMATION_NICHIJI " +
                        ",T_INFORMATION_USERMEI " +
                        ",T_INFORMATION_NAIYO " +
                        "FROM T_INFORMATION " +
                        "ORDER BY T_INFORMATION_NICHIJI DESC";
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(ListData);
                }
                Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / 5)).ToString();
                Paging_now.Text = (1).ToString();
                set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
                set_data(int.Parse(Paging_now.Text));
            }
            catch (Exception)
            {

            }
        }

        private void set_data(int pagenum)
        {
            int viewnum = 5;
            int startrow = (pagenum - 1) * viewnum;
            int addnum = ListData.Rows.Count - startrow;
            if (addnum > viewnum)
            {
                addnum = viewnum;
            }
            tableLayoutPanel5.RowStyles[2] = new RowStyle(SizeType.AutoSize);
            for (int r = 0; r < 5; r++)
            {
                if (r < addnum)
                {
                    tableLayoutPanel6.RowStyles[r] = new RowStyle(SizeType.AutoSize);
                    ((Label)tableLayoutPanel6.Controls["Date" + (r + 1)]).Font = new Font("ＭＳ Ｐゴシック", 9);
                    ((Label)tableLayoutPanel6.Controls["User" + (r + 1)]).Font = new Font("ＭＳ Ｐゴシック", 9);
                    ((Label)tableLayoutPanel6.Controls["Contents" + (r + 1)]).Font = new Font("ＭＳ Ｐゴシック", 9);
                    ((Label)tableLayoutPanel6.Controls["Date" + (r + 1)]).Text = ListData.Rows[startrow + r]["T_INFORMATION_NICHIJI"].ToString();
                    ((Label)tableLayoutPanel6.Controls["User" + (r + 1)]).Text = ListData.Rows[startrow + r]["T_INFORMATION_USERMEI"].ToString();
                    ((Label)tableLayoutPanel6.Controls["Contents" + (r + 1)]).Text = ListData.Rows[startrow + r]["T_INFORMATION_NAIYO"].ToString();
                }
                else
                {
                    tableLayoutPanel6.RowStyles[r] = new RowStyle(SizeType.Absolute,0);
                    ((Label)tableLayoutPanel6.Controls["Date" + (r + 1)]).Text = "";
                    ((Label)tableLayoutPanel6.Controls["User" + (r + 1)]).Text = "";
                    ((Label)tableLayoutPanel6.Controls["Contents" + (r + 1)]).Text = "";
                }
            }
            tableLayoutPanel5.RowStyles[2] = new RowStyle(SizeType.Percent,100);
            tableLayoutPanel6.Invalidate();
            
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            Paging_now.Text = (1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
            //レイアウトロジックを再開する
            this.ResumeLayout();

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            Paging_now.Text = (int.Parse(Paging_now.Text) - 1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
            //レイアウトロジックを再開する
            this.ResumeLayout();

        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            Paging_now.Text = (int.Parse(Paging_now.Text) + 1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
            //レイアウトロジックを再開する
            this.ResumeLayout();

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            Paging_now.Text = (int.Parse(Paging_all.Text)).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
            //レイアウトロジックを再開する
            this.ResumeLayout();

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

        private void button2_Click(object sender, EventArgs e)
        {
            if (GlobalMethod.GetFormsShow("Entry"))
            {
                int PhaseFLG = GlobalMethod.GetIntroductionPhase();
                if (PhaseFLG <= 1)
                {
                    Entry_keikaku_Search form = new Entry_keikaku_Search();
                    form.UserInfos = this.UserInfos;
                    form.Show();
                }
                else
                {
                    Entry_Search form = new Entry_Search();
                    form.UserInfos = this.UserInfos;
                    form.Show();
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 特調野郎系は1画面のみしか開けない制御を入れる
            if (GlobalMethod.GetFormsShow("Tokuchoyaro")) 
            {
                Form f = null;
                // 閉じてたのを全部閉じる（後ろから閉じていく）
                for (int i = System.Windows.Forms.Application.OpenForms.Count - 1; i > 0; i--)
                {
                    f = System.Windows.Forms.Application.OpenForms[i];
                    if (f.Text.IndexOf("特調野郎") >= 0 || f.Text.IndexOf("窓口ミハル") >= 0 || f.Text.IndexOf("特命課長") >= 0 || f.Text.IndexOf("自分大臣") >= 0)
                    {
                        f.Close();
                    }
                }

                Tokuchoyaro form = new Tokuchoyaro();
                form.UserInfos = this.UserInfos;
                form.Show();

                //Tokuchoyaro form = new Tokuchoyaro();
                //form.UserInfos = this.UserInfos;
                //form.Show();
            }
        }

        // 業務日報ボタン
        private void button3_Click(object sender, EventArgs e)
        {
            //Test form = new Test();
            //form.Show();

            string command = "/C " + GlobalMethod.GetCommonValue1("GYOUMU_URL");
            if(command != null)
            {
                //System.Diagnostics.Process.Start("cmd.exe", command);

                System.Diagnostics.ProcessStartInfo psInfo = new System.Diagnostics.ProcessStartInfo("cmd.exe", command);
                psInfo.CreateNoWindow = true; // コンソール・ウィンドウを開かない
                psInfo.UseShellExecute = false; // シェル機能を使用しない

                System.Diagnostics.Process.Start(psInfo);

            }
            else
            {

            }

        }

        // 単価契約
        private void button1_Click(object sender, EventArgs e)
        {
            if (GlobalMethod.GetFormsShow("tanka"))
            {
                tanka form = new tanka();
                form.UserInfos = UserInfos;
                form.Show();
            }
        }


        private void TopMenu_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (((Login)this.Owner).logoff == false)
            {
                CloseForms();
                // アプリケーション（自分自身）を終了させる
                Application.Exit();
            }
        }

        // ログアウト
        private void button5_Click(object sender, EventArgs e)
        {
            ((Login)this.Owner).logoff = true;
            this.Owner.Show();
            CloseForms();

            this.Close();

        }

        private void CloseForms()
        {
            Boolean flg = true;
            while (flg)
            {
                flg = false;
                for (int i = 0; i < Application.OpenForms.Count; i++)
                {
                    Form f = Application.OpenForms[i];

                    if (f.Text.IndexOf("メニュー") == -1 && f.Text.IndexOf("ログイン") == -1)
                    {
                        flg = true;
                        f.Close();
                    }
                }
            }

        }
    }
}
