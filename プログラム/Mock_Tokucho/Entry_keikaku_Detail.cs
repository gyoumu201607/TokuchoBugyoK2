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
                //エントリ君STEP1 KeikakuNendoKurikoshiGakuK ～ KeikakuKaishiNendoまで追加
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
                    ,KeikakuNendoKurikoshiGakuK,KeikakuNendoKurikoshiGakuJ,KeikakuNendoKurikoshiGakuJF,KeikakuNendoKurikoshiGaku
                    ,KeikakuKoukiShuryoubi,KeikakuKoukiKaishibi,KeikakuKaishiNendo
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
            //エントリ君STEP1
            item1_11.Text = ListData.Rows[0][42].ToString() + "年度";
            DateTime dt = new DateTime();
            if(DateTime.TryParse(ListData.Rows[0][41].ToString(),out dt)){
                item1_12.Text = dt.ToString("yyyy/MM/dd");
            }
            if (DateTime.TryParse(ListData.Rows[0][40].ToString(), out dt))
            {
                item1_13.Text = dt.ToString("yyyy/MM/dd");
            }

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
            //エントリ君STEP1
            item2_16.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][39]));
            item2_17.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][38]));
            item2_18.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][37]));
            item2_19.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][36]));
            item2_20.Text = string.Format("{0:C}", Math.Floor((decimal)ListData.Rows[0][39] + (decimal)ListData.Rows[0][38] + (decimal)ListData.Rows[0][37] + (decimal)ListData.Rows[0][36]));

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

        //エントリ君STEP1
        private void btnAnkenCopy_Click(object sender, EventArgs e)
        {
            
            //前年度案件番号（item1_8）がない（空）場合はメッセージ表示して終了
            if (item1_8.Text.Trim() == "")
            {
                MessageBox.Show(GlobalMethod.GetMessage("E11001", ""));
                return;
            }
            else
            {
                //MessageBox.Show("実装中です。しばらくお待ちください。");
                //return;

                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                var dt = new DataTable();
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    //案件情報からデータを取得し、AnkenIDを決定する。
                    //仕様
                    //・KeikakuJouhouの KeikakuZenkaiAnkenBangouとKeikakuZenkaiJutakuBangouがある場合は、KeikakuZenkaiJutakuBangou＋AnkenSaishinFlg＝1で。
                    //　複数の場合は、若番。
                    //　ない場合は、以下
                    //・Ankenjouhouを直に検索されると、主に3種類あります。
                    // ①パターン1 ：　業務日報用9999の案件番号 →　業務日報用なので無視して下さい。計画情報には入りません。
                    // ②パターン2 ：　案件番号に複数の受託番号枝番であるのもの（枝番が01、02～）　→　01 + AnkenSaishinFlg＝1のものを採用して下さい。
                    // ③パターン3 ：　旧案件番行のもの →　若番を採用 + AnkenSaishinFlg＝1。案件番号は現行通り自動採番。

                    //配分率を取得する条件
                    //前回案件番号が受託の場合、業務配分は契約時の業務配分を設定します。未受託の場合は、事前打診・入札の業務配分を設定します。
                    //契約時の業務配分 = [GyoumuHaibun]テーブルの[GyoumuHibunKubun]が30のもの
                    //事前打診・入札の業務配分 = [GyoumuHaibun]テーブルの[GyoumuHibunKubun]が20のもの
                    //　→これらは、案件画面の方で実装する

                    //前回案件番号が受託か否かの判定は下記の通り
                    //・コピー時の受託の状況は、入札画面の応札者一覧で、建設物価調査会様が受託にチェックが入っている場合。
                    //　テーブル ：　NyuusatsuJouhouOusatsusha
                    //企業コード　：　1001 建設物価調査会
                    //受託フラグ　：　NyuusatsuRakusatsuJokyou
                    //　→これらは、案件画面の方で実装する

                    //前年度受託番号がある
                    if (item1_9.Text.Trim() != "")
                    {
                        cmd.CommandText = "SELECT " +
                            "AnkenJouhouID " +
                            "FROM AnkenJouhou " +
                            "WHERE AnkenJutakuBangou = '" + item1_9.Text.Trim() + "' " +
                            "AND AnkenSaishinFlg = 1 " +
                            "ORDER BY AnkenJouhouID DESC";  //若番（新しい方）
                    }
                    else
                    {
                        //
                        cmd.CommandText = "SELECT " +
                            "AnkenJouhouID " +
                            "FROM AnkenJouhou " +
                            "WHERE AnkenAnkenBangou = '" + item1_8.Text.Trim() + "' " +
                            "AND AnkenSaishinFlg = 1 " +
                            "ORDER BY AnkenJutakuBangouEda ASC , AnkenJouhouID DESC";  //枝番=01が優先　NULLまたは空っぽ 若番（新しい方）
                    }
                    

                    var sda = new SqlDataAdapter(cmd);

                    dt.Clear();
                    sda.Fill(dt);

                    //エントリ君修正STEP1　追加要件不具合No5　暫定修正　確定したら、下のコメントアウト部分削除
                    //データが空っぽだったらエラー
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("案件情報が見つかりません。");
                        return;
                    }
                    String AnkenID;
                    //TOPのデータのAnkenIDを渡す。
                    AnkenID = dt.Rows[0][0].ToString();
                    Entry_Input form = new Entry_Input();
                    form.KeikakuID = KeikakuID;
                    form.mode = "keikaku";
                    form.UserInfos = this.UserInfos;
                    form.CopyMode = "1";
                    form.AnkenID = AnkenID;
                    //計画画面から案件番号でコピーする場合のフラグ
                    form.isKeikakuAnkenNew = true;

                    this.Hide();
                    form.ShowDialog(this);
                    //計画情報のカウント取得する
                    cmd.CommandText = "SELECT KeikakuAnkensu FROM KeikakuJouhou WHERE KeikakuID = " + KeikakuID;
                    sda = new SqlDataAdapter(cmd);
                    dt.Clear();
                    sda.Fill(dt);
                    string KeikakuAnkensu = dt.Rows[0][1].ToString();
                    if (KeikakuAnkensu == "")
                    {
                        KeikakuAnkensu = "0";
                    }
                    item1_10.Text = KeikakuAnkensu;
                    if (item1_10.Text == "0")
                    {
                        button3.Enabled = false;
                        button3.BackColor = Color.DarkGray;
                    }
                    else
                    {
                        button3.Enabled = true;
                        //42, 78, 122
                        int R = 42;
                        int G = 78;
                        int B = 122;

                        Color color = Color.FromArgb(R, G, B);
                        button3.BackColor = color;
                    }

                }

                //エントリ君修正STEP1　追加要件不具合No5　暫定修正　確定したらコメント消すこと。
                ////データが空っぽだったらエラー
                //if (dt.Rows.Count == 0)
                //{
                //    MessageBox.Show("案件情報が見つかりません。");
                //    return;
                //}
                //String AnkenID;
                ////TOPのデータのAnkenIDを渡す。
                //AnkenID = dt.Rows[0][0].ToString();
                //Entry_Input form = new Entry_Input();
                //form.KeikakuID = KeikakuID;
                //form.mode = "keikaku";
                //form.UserInfos = this.UserInfos;
                //form.CopyMode = "1";
                //form.AnkenID = AnkenID;
                ////計画画面から案件番号でコピーする場合のフラグ
                //form.isKeikakuAnkenNew = true;

                //this.Hide();
                //form.Show(this);

            }
        }
    }
}
