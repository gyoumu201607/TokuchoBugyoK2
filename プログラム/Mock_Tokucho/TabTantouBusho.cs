using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;

namespace TokuchoBugyoK2
{
    public partial class TabTantouBusho : UserControl
    {
        private string Message = "";
        //private DataTable DT_TantouBusho = new DataTable();
        private DataTable DT_MadoguchiL1Chou = new DataTable();
        private DataTable DT_GaroonTsuikaAtesaki = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();

        public string[] UserInfos;
        public string GamenMode = "";
        public string MadoguchiID = "";
        public string p_Message = "";


        public event EventHandler TabTantouBushoMessage;

        // 親側にメッセージを表示
        private void OnTabTantouBushoMessage()
        {
            if (TabTantouBushoMessage != null)
            {
                TabTantouBushoMessage(this, EventArgs.Empty);
            }
        }

        public void TabTabTantouBushoGetData()
        {
            get_data(2);
        }

        public TabTantouBusho()
        {
            InitializeComponent();
        }

        // 担当部所タブの更新ボタン
        private void button_update_2_Click(object sender, EventArgs e)
        {
            // エラーフラグ true：エラー false：正常
            Boolean errorFlg = false;
            // エラーメッセージフラグ
            Boolean messageFlg1 = true;
            Boolean messageFlg2 = true;

            string tantousha = "";

            //set_error("", 0);
            // エラーチェック
            for (int i = 1; i < c1FlexGrid5.Rows.Count; i++)
            {
                // 担当者が重複していないか確認
                if (c1FlexGrid5.Rows[i][3] != null && !"".Equals(c1FlexGrid5.Rows[i][3].ToString()) && !"0".Equals(c1FlexGrid5.Rows[i][3].ToString()) && !"".Equals(tantousha))
                {
                    if (tantousha.IndexOf(c1FlexGrid5.Rows[i][3].ToString()) > -1)
                    {
                        messageFlg2 = true;
                        errorFlg = true;
                    }
                }

                // 部所が選択されており、担当者が空の場合、エラー
                // 2:担当部所 3:担当者
                if (c1FlexGrid5.Rows[i][2] != null && c1FlexGrid5.Rows[i][3] != null && !"".Equals(c1FlexGrid5.Rows[i][2].ToString()) && !"0".Equals(c1FlexGrid5.Rows[i][2].ToString()) && "".Equals(c1FlexGrid5.Rows[i][3].ToString()))
                {
                    messageFlg1 = true;
                    errorFlg = true;
                }

                // 担当者が空でない場合、変数に格納する
                if (c1FlexGrid5.Rows[i][3] != null && !"".Equals(c1FlexGrid5.Rows[i][3].ToString()) && !"0".Equals(c1FlexGrid5.Rows[i][3].ToString()))
                {
                    tantousha += c1FlexGrid5.Rows[i][3] + ",";
                }
            }
        }

        // 担当部所 行追加
        private void button_GaroonAtesakiGridAdd_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid5.BeginUpdate();

            c1FlexGrid5.Rows.Add();
            Resize_Grid("c1FlexGrid5");

            //描画再開
            c1FlexGrid5.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();

        }

        private void get_data(int tab)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            try
            {
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();

                    //調査担当者の取得
                    cmd.CommandText = "SELECT " +
                        "CASE  " +
                            "WHEN MadoguchiHoukokuzumi = 1 THEN 8 " +
                            //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShinchokuJoukyou = 6 THEN 6 " +　//MadoguchiShinchokuJoukyouが2桁コードに変更
                            //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShinchokuJoukyou = 3 THEN 5 " +
                            //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<GETDATE() THEN 1 " +
                            //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<GETDATE() +3 AND MadoguchiShinchokuJoukyou<> 3 THEN 2 " +
                            //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<GETDATE() +7 AND MadoguchiShinchokuJoukyou<> 3 THEN 3 " +

                            //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShinchokuJoukyou = 80 THEN 6 " +
                            //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShinchokuJoukyou = 70 THEN 5 " +
                            //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<'" + DateTime.Today + "' THEN 1 " +
                            //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<='" + DateTime.Today.AddDays(3) + "' AND MadoguchiShinchokuJoukyou<> 70 THEN 2 " +
                            //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<='" + DateTime.Today.AddDays(7) + "' AND MadoguchiShinchokuJoukyou<> 70 THEN 3 " +

                            "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShinchoku = 80 THEN 6 " +
                            "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShinchoku = 70 THEN 5 " +
                            "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShimekiribi<'" + DateTime.Today + "' THEN 1 " +
                            "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShimekiribi<='" + DateTime.Today.AddDays(3) + "' AND MadoguchiL1ChousaShinchoku <> 70 THEN 2 " +
                            "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShimekiribi<='" + DateTime.Today.AddDays(7) + "' AND MadoguchiL1ChousaShinchoku <> 70 THEN 3 " +
                            "ELSE 4 " +
                        "END AS 'SortID'  " +
                        ", MadoguchiL1ChousaCD " +
                        ", MadoguchiL1ChousaBushoCD " +
                        ", MadoguchiL1ChousaTantoushaCD " +
                        ", MadoguchiL1ChousaShimekiribi " +
                        ", MadoguchiL1ChousaShinchoku " +
                        ", MadoguchiL1ChousaKakunin " +
                        ", BushokanriboKamei " +
                        "FROM MadoguchiJouhouMadoguchiL1Chou " +
                        "LEFT JOIN MadoguchiJouhou ON MadoguchiJouhouMadoguchiL1Chou.MadoguchiID = MadoguchiJouhou.MadoguchiID " +
                        "LEFT JOIN Mst_Busho ON MadoguchiJouhouMadoguchiL1Chou.MadoguchiL1ChousaBushoCD = Mst_Busho.GyoumuBushoCD " +
                        "WHERE MadoguchiJouhou.MadoguchiID = '" + MadoguchiID + "' ORDER BY SortID ";

                    Console.WriteLine(cmd.CommandText);
                    var sda = new SqlDataAdapter(cmd);
                    DT_MadoguchiL1Chou.Clear();
                    sda.Fill(DT_MadoguchiL1Chou);

                    //Garoon追加宛先の取得
                    cmd.CommandText = "SELECT " +
                        "  GaroonTsuikaAtesakiID " +
                        ", GaroonTsuikaAtesakiBushoCD " +
                        ", GaroonTsuikaAtesakiTantoushaCD " +
                        "FROM GaroonTsuikaAtesaki " +
                        "WHERE GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "' AND GaroonTsuikaAtesakiDeleteFlag <> 1 ";

                    Console.WriteLine(cmd.CommandText);
                    sda = new SqlDataAdapter(cmd);
                    DT_GaroonTsuikaAtesaki.Clear();
                    sda.Fill(DT_GaroonTsuikaAtesaki);

                }

            }
            catch (Exception)
            {
                throw;
            }
            set_data(tab);

            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void set_data(int tab)
        {
            if (tab == 2)
            {
                //担当部所の調査担当者を初期化
                c1FlexGrid1.Rows.Count = 1;
                //c1FlexGrid1.Rows.Count = 31;
                //担当部所のGaroon追加宛先を初期化
                c1FlexGrid5.Rows.Count = 1;

                if (DT_MadoguchiL1Chou != null)
                {
                    //描画停止
                    c1FlexGrid1.BeginUpdate();
                    //担当部署の部所一覧を初期化
                    for (int k = 1; k < 26; k++)
                    {
                        ((CheckBox)tableLayoutPanel17.Controls["KyoroykuBusho" + k]).Checked = false;
                    }
                    for (int i = 0; i < DT_MadoguchiL1Chou.Rows.Count; i++)
                    {
                        //調査担当者をセット
                        //// ヘッダーで1行使っているので、31行目以降追加する
                        //if (i >= 30)
                        //{
                        //    c1FlexGrid1.Rows.Add();
                        //}
                        c1FlexGrid1.Rows.Add();

                        for (int k = 0; k < c1FlexGrid1.Cols.Count; k++)
                        {
                            c1FlexGrid1.Rows[i + 1][k] = DT_MadoguchiL1Chou.Rows[i][k];
                        }
                        //部所一覧をセット
                        for (int k = 1; k < 26; k++)
                        {
                            if (((CheckBox)tableLayoutPanel17.Controls["KyoroykuBusho" + k]).Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString()))
                            {
                                ((CheckBox)tableLayoutPanel17.Controls["KyoroykuBusho" + k]).Checked = true;
                            }
                        }
                    }
                    //描画再開
                    c1FlexGrid1.EndUpdate();
                }
                if (DT_GaroonTsuikaAtesaki != null)
                {
                    // データカウント
                    int count = 0;
                    //描画停止
                    c1FlexGrid5.BeginUpdate();
                    for (int i = 0; i < DT_GaroonTsuikaAtesaki.Rows.Count; i++)
                    {
                        count += 1;
                        //調査担当者をセット
                        //// ヘッダーで1行使っているので、31行目以降追加する
                        //if (i >= 30) { 
                        //c1FlexGrid5.Rows.Add();
                        //}
                        c1FlexGrid5.Rows.Add();
                        for (int k = 1; k < c1FlexGrid5.Cols.Count; k++)
                        {
                            c1FlexGrid5.Rows[i + 1][k] = DT_GaroonTsuikaAtesaki.Rows[i][k - 1];
                        }
                    }

                    //描画再開
                    c1FlexGrid5.EndUpdate();
                }
                Resize_Grid("c1FlexGrid1");
                Resize_Grid("c1FlexGrid5");
            }
        }

        // Gridのサイズ調整
        public void Resize_Grid(string name)
        {
            Control[] cs;
            cs = this.Controls.Find(name, true);
            if (cs.Length > 0)
            {
                var fx = (C1.Win.C1FlexGrid.C1FlexGrid)cs[0];

                int rowcount = fx.Rows.Count;

                int h = 0;
                for (int i = 0; i < fx.Rows.Count; i++)
                {
                    if (fx.Rows[i].Height == -1)
                    {
                        h += 22;
                    }
                    else
                    {
                        h += fx.Rows[i].Height;
                    }
                }
                fx.Height = 4 + h;
                //c1FlexGrid1.Height = 4 + h;

                int w = 0;
                for (int i = 0; i < fx.Cols.Count; i++)
                {
                    if (fx.Cols[i].Width == -1)
                    {
                        w += 100;
                    }
                    else
                    {
                        w += fx.Cols[i].Width;
                    }
                }
                if (fx.Width < 4 + w)
                {
                    fx.Height += 18;
                }
            }
        }

        //担当部所タブのGaroon追加宛先Gridをクリック時
        private void c1FlexGrid5_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid5.HitTest(new Point(e.X, e.Y));
            //担当者列をクリック
            if (hti.Row > 0 && hti.Column == 3)
            {
                Popup_ChousainList form = new Popup_ChousainList();
                if (c1FlexGrid5[hti.Row, 2] != null && c1FlexGrid5[hti.Row, 2].ToString() != "")
                {
                    form.Busho = c1FlexGrid5[hti.Row, 2].ToString();
                }
                else
                {
                    //form.Busho = item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString();
                }
                form.program = "madoguchi";
                //if (item1_MadoguchiTourokuNendo.Text != "")
                //{
                //    form.nendo = item1_MadoguchiTourokuNendo.SelectedValue.ToString();
                //}
                //form.ShowDialog();
                
                if (form.ReturnValue != null && form.ReturnValue[0] != null)
                {
                    c1FlexGrid5[hti.Row, 2] = form.ReturnValue[2];
                    c1FlexGrid5[hti.Row, 3] = form.ReturnValue[0];
                }
            }
            //削除列をクリック
            if (hti.Row > 0 && hti.Column == 0)
            {
                //if (GlobalMethod.outputMessage("I10002", "", 1) == DialogResult.OK)
                //{
                //    c1FlexGrid5.Rows.Remove(hti.Row);
                //    Resize_Grid("c1FlexGrid5");
                //}
            }
        }
    }
}
