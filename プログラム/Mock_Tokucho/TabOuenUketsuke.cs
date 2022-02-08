using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Drawing;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TokuchoBugyoK2
{
    public partial class TabOuenUketsuke : UserControl
    {
        private string Message = "";
        private DataTable DT_Ouenuketsuke = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();

        public string[] UserInfos;
        public string GamenMode = "";
        public string MadoguchiID = "";
        public string p_Message = "";
        public string MadoguchiChousaKubunShibuHonbu = "";

        // 
        public event EventHandler TabOuenUketsukeMessage;

        public void TabOuenUketsukeGetData()
        {
            get_data(5);

        }

        // 親側にメッセージを表示
        private void OnTabOuenUketsukeMessage()
        {
            if (TabOuenUketsukeMessage != null)
            {
                TabOuenUketsukeMessage(this, EventArgs.Empty);
            }
        }

        public TabOuenUketsuke()
        {
            InitializeComponent();
        }

        private void TabOuenUketsuke_Load(object sender, EventArgs e)
        {
            //get_data(5);
        }

        // 応援受付の依頼書出力ボタン
        private void btn6_Iraisho(object sender, EventArgs e)
        {
            // 支→本のみ表示
            if ("1".Equals(MadoguchiChousaKubunShibuHonbu))
            {
                UketsukeIcon.Visible = true;
            }
            else
            {
                UketsukeIcon.Visible = false;
            }
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

                    //窓口情報取得
                    cmd.CommandText = "SELECT " +
                        "OuenKanriNo " +                     // 0:管理番号
                        ",OuenJoukyou " +                    // 1:応援状況
                        ",OuenUketsukeDate " +               // 2:応援受付日
                        ",OuenKanryou " +                    // 3:応援完了
                        ",OuenHoukokuJishibi " +             // 4:応援完了日
                        ",MadoguchiChousaKubunShibuHonbu " + // 5:支→本
                        "FROM OuenUketsuke ou " +
                        "LEFT JOIN MadoguchiJouhou mj ON mj.MadoguchiID = ou.MadoguchiID " + 
                        "WHERE ou.MadoguchiID = " + MadoguchiID + " AND OuenDeleteFlag != 1 ";

                    var sda = new SqlDataAdapter(cmd);
                    DT_Ouenuketsuke.Clear();
                    sda.Fill(DT_Ouenuketsuke);

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
            // 応援受付
            if (tab == 5 && DT_Ouenuketsuke != null && DT_Ouenuketsuke.Rows.Count > 0)
            {
                // 管理番号
                item5_Kanribangou.Text = DT_Ouenuketsuke.Rows[0][0].ToString();
                // 調査区分 支→本
                MadoguchiChousaKubunShibuHonbu = DT_Ouenuketsuke.Rows[0][5].ToString();

                // 応援状況
                if (DT_Ouenuketsuke.Rows[0][1] != null && DT_Ouenuketsuke.Rows[0][1].ToString() == "1")
                {
                    item5_UketsukeJoukyo.Checked = true;
                    // 支→本のみ表示
                    if ("1".Equals(MadoguchiChousaKubunShibuHonbu))
                    {
                        UketsukeIcon.Visible = true;

                        // チェック時は完了アイコン
                        UketsukeIcon.Image = Image.FromFile("Resource/kan.png");
                    }
                    else
                    {
                        UketsukeIcon.Visible = false;
                    }
                }
                else
                {
                    UketsukeIcon.Visible = false;
                }
                // 応援受付日
                if (DT_Ouenuketsuke.Rows[0][2] != null && DT_Ouenuketsuke.Rows[0][2].ToString() != "")
                {
                    item5_UketsukeDate.Text = DT_Ouenuketsuke.Rows[0][2].ToString();
                    item5_UketsukeDate.CustomFormat = "";
                }
                else
                {
                    item5_UketsukeDate.CustomFormat = " ";
                }
                // 応援完了
                if (DT_Ouenuketsuke.Rows[0][3] != null && DT_Ouenuketsuke.Rows[0][3].ToString() == "1")
                {
                    item5_OuenKanryo.Checked = true;
                    // 支→本のみ表示
                    if ("1".Equals(MadoguchiChousaKubunShibuHonbu))
                    {
                        KanryouIcon.Visible = true;
                    }
                    else
                    {
                        KanryouIcon.Visible = false;
                    }
                }
                else
                {
                    KanryouIcon.Visible = false;
                }
                // 応援完了日
                if (DT_Ouenuketsuke.Rows[0][4] != null && DT_Ouenuketsuke.Rows[0][4].ToString() != "")
                {
                    item5_OuenKanryoDate.Text = DT_Ouenuketsuke.Rows[0][4].ToString();
                    item5_OuenKanryoDate.CustomFormat = "";
                }
                else
                {
                    item5_OuenKanryoDate.CustomFormat = " ";
                }
            }
        }

        // 更新ボタン
        private void btnOuenuketsukeUpdate_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("更新を行いますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                UpdateMadoguchi(5);
            }
        }


        private void UpdateMadoguchi(int tab)
        {
            //更新成功フラグ
            Boolean UpdateFlag = true;

            // 応援受付タブ
            if (tab == 5)
            {
                string[,] SQLData = new string[1, 5];
                //SQLData[0, 0] = item5_Kanribangou.Text; // 管理番号は調査概要でやっているので除外
                // SQLData[0, 0]:応援状況
                // SQLData[0, 1]:応援受付日
                // SQLData[0, 2]:応援完了
                // SQLData[0, 3]:応援完了日
                if (item5_UketsukeJoukyo.Checked == true)
                {
                    SQLData[0, 0] = "1";
                }
                if (item5_UketsukeDate.CustomFormat == "")
                {
                    SQLData[0, 1] = item5_UketsukeDate.Text;
                }
                if (item5_OuenKanryo.Checked == true)
                {
                    SQLData[0, 2] = "1";
                }
                if (item5_OuenKanryoDate.CustomFormat == "")
                {
                    SQLData[0, 3] = item5_OuenKanryoDate.Text;
                }
                string mes = "";
                //GlobalMethod.MadoguchiUpdate_SQL(tab, MadoguchiID, SQLData, out mes, UserInfos);

                //set_error("", 0);
                //set_error(mes);

                p_Message = mes;
                // 親にメッセージを表示
                OnTabOuenUketsukeMessage();
            }
            if (UpdateFlag)
            {
                get_data(tab);
            }
        }

        // 応援受付タブ 応援状況チェック
        private void Uketsuke_CheckStateChange(object sender, EventArgs e)
        {
            //チェック状態の確認
            switch (((CheckBox)(sender)).CheckState)
            {
                case CheckState.Checked:
                    // チェック
                    item5_UketsukeDate.Value = System.DateTime.Today;
                    item5_UketsukeDate.CustomFormat = "";

                    if ("1".Equals(MadoguchiChousaKubunShibuHonbu))
                    {
                        UketsukeIcon.Visible = true;
                    }
                    // チェック時は完了アイコン
                    UketsukeIcon.Image = Image.FromFile("Resource/kan.png");
                    break;
                case CheckState.Unchecked:
                    // 未チェック
                    item5_UketsukeDate.Value = System.DateTime.Today;
                    item5_UketsukeDate.CustomFormat = " ";
                    UketsukeIcon.Visible = false;
                    break;
            }
        }

        // 応援完了
        private void OuenKanryou_CheckStateChange(object sender, EventArgs e)
        {
            //チェック状態の確認
            switch (((CheckBox)(sender)).CheckState)
            {
                case CheckState.Checked:
                    // チェック
                    item5_OuenKanryoDate.Value = System.DateTime.Today;
                    item5_OuenKanryoDate.CustomFormat = "";
                    KanryouIcon.Visible = true;
                    break;
                case CheckState.Unchecked:
                    // 未チェック
                    item5_OuenKanryoDate.Value = System.DateTime.Today;
                    item5_OuenKanryoDate.CustomFormat = " ";
                    KanryouIcon.Visible = false;
                    break;
            }
        }

        private void dateTimePicker_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Delete)
            {
                ((DateTimePicker)sender).CustomFormat = " ";
            }
        }

        private void dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            ((DateTimePicker)sender).CustomFormat = "";

            DateTime dt = ((DateTimePicker)sender).Value;
            ((DateTimePicker)sender).Text = dt.ToString("yyyy/MM/dd");
        }
    }
}
