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
using System.IO;


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

        //報告書改善対応：帳票更新APL関連
        private const string APL_VERSION_CHOUHYOUUPDATE = "APL_VERSION_CHOUHYOUUPDATE";
        private const char LINE_DATA_EQUAL = '=';
        private const string ENV_USERNAME = "%username%";
        //帳票
        private enum RESULT_CODE
        {
            OK = 0
            ,NG = 1
        }

        /// <summary>
        /// 帳票出力アプリケーション（KMJ社作成）の最新版をコピーするためのメソッド
        /// </summary>
        private void ReportExeUpdate()
        {
            //ローディング画面
            Popup_Loading Loading = new Popup_Loading();
            Loading.StartPosition = FormStartPosition.CenterScreen;
            //Loading.Show();

            try
            {
 
                //環境変数からログインユーザ名を取得する
                string loginUser = System.Environment.ExpandEnvironmentVariables(ENV_USERNAME);

                //帳票更新APLのバージョンをDBから取得する。
                string latestVersion = null2string(GlobalMethod.GetCommonValue1("APL_VERSION_CHOUHYOUUPDATE"));
                //帳票更新APLのサーバのフォルダを取得   フォルダの情報はすべて末尾に¥がついてる想定
                string serverReportAplFolder = null2string(GlobalMethod.GetCommonValue1("CHOUHYOUUPDATE_EXE_SERVER_FOLDER"));

                //バージョンやサーバパスがとれなかったらエラーで何もしなくて良いか。
                if (latestVersion == "" || serverReportAplFolder == "")
                {
                    return;

                }

                //帳票更新APLを置くフォルダを取得
                string repAplFolder = null2string(GlobalMethod.GetCommonValue1("CHOUHYOUUPDATE_EXE_FOLDER")).Replace(ENV_USERNAME, loginUser);

                //configファイルのパスを生成
                string configFilePath = repAplFolder + null2string(GlobalMethod.GetCommonValue1("CHOUHYOUUPDATE_CLIENT_CONFIG_FILE"));

                //帳票更新APLファイルのパスを生成
                string reportAplName = null2string(GlobalMethod.GetCommonValue1("CHOUHYOUUPDATE_EXE_NAME"));
                string localReportAplPath = repAplFolder + reportAplName;
                string ServerReportAplPath = serverReportAplFolder + reportAplName;
                //帳票更新APLがDB接続ファイルが必要になったらしい。
                string reportAplDbConfigName = null2string(GlobalMethod.GetCommonValue1("CHOUHYOUUPDATE_EXE_NAME","2"));
                string localReportAplDbConfigPath = repAplFolder + reportAplDbConfigName;
                string ServerReportAplDbConfigPath = serverReportAplFolder + reportAplDbConfigName;

                bool isReportAplCopy = false;

                //configファイルが存在していた
                if (isExistsFile(configFilePath))
                {
                    //configファイルを開き、バージョン番号を取得する。
                    string localVersion = getReportAplVersion(configFilePath);
                    //ファイルから取得したバージョン番号と、DBから取得したバージョン番号が異なる場合、帳票更新アプリをコピーするフラグを立てる
                    if (latestVersion != localVersion)
                    {
                        isReportAplCopy = true;
                    }

                }
                else //configファイルが存在していなかった。
                {
                    //帳票更新APLを置くフォルダを作成する。mkdir
                    createDir(repAplFolder);

                    //帳票更新APLをサーバからローカルにコピーするフラグを立てる
                    isReportAplCopy = true;
                }

                //帳票更新APLのコピーフラグが立っている場合、コピー実施および、Exe起動する。
                if (isReportAplCopy)
                {
                    //帳票サーバから帳票更新APLをコピーする。
                    //サーバのファイルの存在確認
                    if (isExistsFile(ServerReportAplPath) && isExistsFile(ServerReportAplDbConfigPath))
                    {
                        Loading.Show();

                        //先にDBConfigファイルを取得
                        File.Copy(ServerReportAplDbConfigPath, localReportAplDbConfigPath, true);
                        //最新の帳票更新APLをローカルPCにコピー
                        File.Copy(ServerReportAplPath, localReportAplPath, true);

                        //最新の帳票更新APLを起動する。
                        //成功するまでリトライするカウント
                        const int RETRY_COUNT= 1;
                        bool isSuccess = false;

                        //最新の帳票更新APLを起動する。
                        for (int i = 0; i <= RETRY_COUNT; i++)
                        {
                            //2回目以降は1秒スリープする
                            if (i > 0)
                            {
                                System.Threading.Thread.Sleep(1000);
                            }
                            //帳票更新APLを終了待ちで起動する。最新バージョン番号をパラメータとして起動
                            System.Diagnostics.ProcessStartInfo psInfo = new System.Diagnostics.ProcessStartInfo();
                            psInfo.FileName = localReportAplPath;
                            psInfo.Arguments= latestVersion;
                            psInfo.CreateNoWindow = true; // コンソール・ウィンドウを開かない
                            psInfo.UseShellExecute = false; // シェル機能を使用しない

                            System.Diagnostics.Process p = System.Diagnostics.Process.Start(psInfo);
                            p.WaitForExit();
                            //成功したら終了
                            if(p.ExitCode == (int)RESULT_CODE.OK)
                            {
                                Console.WriteLine("Result Code = " + p.ExitCode.ToString());
                                isSuccess = true;
                                break;
                            }
                            
                        }
                        //成功
                        if (isSuccess)
                        {
                            //configファイルのバージョンを作成する、または書き換える
                            writeFileExecResult(configFilePath, latestVersion);
                        }
                        else　//失敗
                        {
                            Loading.Close();
                            //更新プログラム実行中にエラーが発生しました。
                            MessageBox.Show(GlobalMethod.GetMessage("E00011", ""), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        Loading.Close();
                        //プログラム更新時にエラーが発生しました。
                        MessageBox.Show(GlobalMethod.GetMessage("E00010", ""), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Loading.Close();
                //プログラム更新時にエラーが発生しました。
                MessageBox.Show(GlobalMethod.GetMessage("E00010", ""), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //ロード中ファイル閉じる
                Loading.Close();
            }

        }
        /// <summary>
        /// ディレクトリの存在を確認する
        /// </summary>
        /// <param name="dir">ディレクトリ</param>
        /// <returns>存在有無（true:あり false:なし）</returns>
        private bool isExistsDir(string dir)
        {
            if (System.IO.Directory.Exists(dir))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// ディレクトリを作成する
        /// </summary>
        /// <param name="dir">ディレクトリ</param>
        private void createDir(string dir)
        {
            System.IO.Directory.CreateDirectory(dir);
        }

        /// <summary>
        /// ファイルの存在を確認する
        /// </summary>
        /// <param name="fileName">ファイル名</param>
        /// <returns>存在有無（true:あり false:なし）</returns>
        private bool isExistsFile(string fileName)
        {
            if (System.IO.File.Exists(fileName))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// バージョン番号をconfigファイルに記載する
        /// </summary>
        /// <param name="filePath">configファイルパス</param>
        /// <param name="versionNo">バージョン番号</param>
        private void writeFileExecResult(string filePath , string versionNo)
        {
            //サンプル
            // APL_VERSION_CHOUHYOUUPDATE = 2.0.0.82.01

            string[] _writeStr;
            _writeStr = new string[1];
            _writeStr[0] = APL_VERSION_CHOUHYOUUPDATE + LINE_DATA_EQUAL + versionNo;

            // 文字コード(ここでは、Shift JIS)
            //Encoding enc = Encoding.GetEncoding("UTF-8");
            Encoding enc = new System.Text.UTF8Encoding();

            // ファイルが存在しているときは、上書きする
            System.IO.File.WriteAllLines(filePath, _writeStr, enc);
        }

       ///configファイルを読み出し、バージョン番号を返却する。
       ///ファイルの存在確認は事前に実施すること。
       private string getReportAplVersion(string filePath)
        {
            string versionNo = "";

            foreach (string line in System.IO.File.ReadLines(filePath))
            {
                System.Console.WriteLine(line);
                string[] arrData = line.Split(LINE_DATA_EQUAL);
                if (arrData.Length > 1)
                {
                    if(arrData[0].Trim() == APL_VERSION_CHOUHYOUUPDATE)
                    {
                        versionNo = arrData[1].Trim();
                            break;
                    }
                }
 
            }

            return versionNo;

        }

        private string null2string(string buff)
        {
            if (buff == null)
            {
                return "";
            }
            else
            {
                return buff;
            }
        }

        private void TopMenu_Shown(object sender, EventArgs e)
        {
            //メイン画面の表示をまず行う。
            this.Refresh();
            //報告書改善対応
            ReportExeUpdate();
        }
    }
}
