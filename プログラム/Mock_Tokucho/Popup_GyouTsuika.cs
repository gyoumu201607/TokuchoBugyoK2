using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;

namespace TokuchoBugyoK2
{
    public partial class Popup_GyouTsuika : Form
    {
        public string[] UserInfos;
        public string[] ReturnValue = new string[8];
        GlobalMethod GlobalMethod = new GlobalMethod();

        public int Nendo = DateTime.Today.Year;
        private int ToNendo = 0;
		private string RetireFlg;

		public Popup_GyouTsuika()
        {
            InitializeComponent();

            // コンボボックスにマウスホイールイベントを付与
            this.comboBox_ChousaBusho.MouseWheel += item_MouseWheel;
            this.comboBox_ChousaTantousha.MouseWheel += item_MouseWheel;

        }

        private void Popup_GyouTsuika_Load(object sender, EventArgs e)
        {

            ToNendo = Nendo + 1;

            // コンボボックスの設定
            set_combo();

            // 値の取得
            get_data();

        }

        //コンボボックス設定
        private void set_combo()
        {

            // 調査担当部所 コンボ
            DataTable combodt1 = new System.Data.DataTable();
            string discript = "BushokanriboKamei ";
            string value = "GyoumuBushoCD ";
            string table = "Mst_Busho ";
            string where = "BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != '' "
                         //+ " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " 
                         //+ " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) "
                         + " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) "
                         + " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/1' ) "
                         + " ORDER BY BushoMadoguchiNarabijun "
                         ;
            //コンボボックスデータ取得
            combodt1 = GlobalMethod.getData(discript, value, table, where);
            if (combodt1 != null)
            {
                DataRow dr = combodt1.NewRow();
                combodt1.Rows.InsertAt(dr, 0);
            }
            comboBox_ChousaBusho.DisplayMember = "Discript";
            comboBox_ChousaBusho.ValueMember = "Value";
            comboBox_ChousaBusho.DataSource = combodt1;

        }

        private void get_data()
        {
            //Resize_Grid("c1FlexGrid1");
        }

        private void set_data()
        {
            // 特に処理なし
        }

        // 追加ボタン
        private void btn_LineAdd_Click(object sender, EventArgs e)
        {

            // I20317:調査品目の行を追加しますが宜しいですか？
            if (MessageBox.Show(GlobalMethod.GetMessage("I20317", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                // 画面の値を返す。gridの更新は元画面でさせる。
                ReturnValue[0] = comboBox_ChousaBusho.SelectedValue.ToString();         // 調査担当部所CD
                ReturnValue[1] = comboBox_ChousaBusho.Text;                             // 調査担当部所
                ReturnValue[2] = comboBox_ChousaTantousha.SelectedValue.ToString();     // 調査担当者CD
                ReturnValue[3] = comboBox_ChousaTantousha.Text;                         // 調査担当者
                ReturnValue[4] = textBox_TankaTekiyouChiiki.Text;                       // 単価適用地域
                ReturnValue[5] = textBox_TuikaGyousuu.Text;                             // 追加行数
                ReturnValue[6] = textBox_ZentaiJunKaishiNo.Text;                        // 全体順開始番号

                //調査員の退職フラグ取得
                this.GetUserRetireFlg();
                ReturnValue[7] = this.RetireFlg;                                        // 退職フラグ

                this.Close();
            }

        }

        //調査員の退職フラグ取得
        private void GetUserRetireFlg()
        {
            string kojinCD = comboBox_ChousaTantousha.SelectedValue.ToString();
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();

                // 調査員マスタの値の取得
                var dtprint = new DataTable();
                cmd.CommandText 
= @"SELECT    case 
        when RetireFLG = 0 then '0'
        when RetireFLG = 1 then '1'
    end as RetireFLG "
                                + "  FROM Mst_Chousain"
                                + " WHERE KojinCD = " + kojinCD;
                // データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dtprint);
                if (dtprint.Rows.Count > 0)
                {
                    this.RetireFlg = dtprint.Rows[0]["RetireFLG"].ToString();
                }
                conn.Close();

            }


        }

		// キャンセルボタン
		private void button_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void set_error(string mes, int flg = 1)
        {
            if (flg == 0)
            {
                ErrorMessage.Text = "";
                ErrorBox.Visible = false;
            }
            else
            {
                if (ErrorMessage.Text != "")
                {
                    ErrorMessage.Text += Environment.NewLine;
                }
                ErrorMessage.Text += mes;
                ErrorBox.Visible = true;
            }
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

        // コンボボックス変更時（調査担当部所）
        private void comboBox_ChousaBushoSelectedIndexChanged(object sender, EventArgs e)
        {
            // 調査担当部所を変更することで、調査担当者のコンボを生成する
            // 調査担当者 コンボ
            DataTable combodt2 = new System.Data.DataTable();
            string discript = "ChousainMei ";
            string value = "KojinCD  ";
            string table = "Mst_Chousain ";
            string where = "ChousainDeleteFlag = 0";
            if (comboBox_ChousaBusho.SelectedValue != null && comboBox_ChousaBusho.SelectedValue.ToString() != "")
            {
                where += " AND GyoumuBushoCD = " + comboBox_ChousaBusho.SelectedValue;
            }
            //where += " AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + Nendo + "/4/01' ) "
            //       + " AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' ) "
            where += " AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + ToNendo + "/3/31' ) "
                   + " AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + Nendo + "/4/1' ) "
                   + " ORDER BY KojinCD "
                   ;
            //コンボボックスデータ取得
            combodt2 = GlobalMethod.getData(discript, value, table, where);
            if (combodt2 != null)
            {
                DataRow dr = combodt2.NewRow();
                combodt2.Rows.InsertAt(dr, 0);
            }
            comboBox_ChousaTantousha.DisplayMember = "Discript";
            comboBox_ChousaTantousha.ValueMember = "Value";
            comboBox_ChousaTantousha.DataSource = combodt2;
        }

        // コンボボックス変更時（調査担当者）
        private void comboBox_ChousaTantoushaSelectedIndexChanged(object sender, EventArgs e)
        {
            // 調査担当部所が空で調査担当者が選択されていた時、対象者の部所を自動設定する。
            if (comboBox_ChousaBusho.Text == "" && comboBox_ChousaTantousha.SelectedValue.ToString() != "")
            {
                string kojinCD = comboBox_ChousaTantousha.SelectedValue.ToString();
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();

                    // 調査員マスタの値の取得
                    var dtprint = new DataTable();
                    cmd.CommandText = "SELECT GyoumuBushoCD "
                                    + "  FROM Mst_Chousain"
                                    + " WHERE KojinCD = " + comboBox_ChousaTantousha.SelectedValue;
                    // データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dtprint);
                    if (dtprint.Rows.Count > 0)
                    {
                        comboBox_ChousaBusho.SelectedValue = dtprint.Rows[0]["GyoumuBushoCD"].ToString();
                    }
                    conn.Close();

                    // 調査担当部所を設定した際に調査担当者のコンボが作り直されるため、値をセットし直す。
                    comboBox_ChousaTantousha.SelectedValue = kojinCD;
                }
            }
        }

        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        private void textbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.' && e.KeyChar != '-')
            {
                e.Handled = true;
            }
        }

        private void textBox_ValidatedDecimal(object sender, EventArgs e)
        {
            string tmp = ((System.Windows.Forms.TextBox)sender).Text;
            ((System.Windows.Forms.TextBox)sender).Text = ChangeDecimalText(tmp);
        }

        private string ChangeDecimalText(string str)
        {
            double.TryParse(str.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty), out double num);

            string restr = string.Format("{0:F2}", num);
            return restr;
        }

        private void textBox_ValidatedNumeric(object sender, EventArgs e)
        {
            string tmp = ((System.Windows.Forms.TextBox)sender).Text;
            ((System.Windows.Forms.TextBox)sender).Text = ChangelongText(tmp);
        }

        private string ChangelongText(string str)
        {
            if (str == "")
            {
                str = "0";
            }
            double.TryParse(str.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty), out double num);

            string restr = string.Format("{0}", num);
            return restr;
        }
    }
}
