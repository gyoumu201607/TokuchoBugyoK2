using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TokuchoBugyoK2
{
    public partial class Tokuchoyaro : Form
    {
        public string[] UserInfos;
        GlobalMethod GlobalMethod = new GlobalMethod();
        public Boolean ReSearch = false;

        DataTable DT_UserShinchoku = new DataTable();
        DataTable DT_MadoguchiShinshoku = new DataTable();
        DataTable DT_TantoushaChange = new DataTable();
        DataTable DT_BlankShinchoku = new DataTable();

        public Tokuchoyaro()
        {
            InitializeComponent();
        }


        private void Tokuchoyaro_Load(object sender, EventArgs e)
        {
            // ホイール制御
            this.src_Nendo.MouseWheel += item_MouseWheel; // 売上年度
            this.src_ShuFuku.MouseWheel += item_MouseWheel; // 主副

            label3.Text = UserInfos[3] + "：" + UserInfos[1];
            set_combo();

            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid2.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid2.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid3.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid3.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid4.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid4.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");


        }

        // TOP
        private void button6_Click(object sender, EventArgs e)
        {
            //Tokuchoyaro form = new Tokuchoyaro();
            //form.UserInfos = this.UserInfos;
            //form.Show();
            //this.Close();
            //Form f = null;
            //Boolean openFlg = false;
            //for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
            //{
            //    f = System.Windows.Forms.Application.OpenForms[i];
            //    if (f.Text.IndexOf("特調野郎") >= 0 && f.Text.IndexOf("編集") <= -1)
            //    {
            //        f.Show();
            //        openFlg = true;
            //        break;
            //    }
            //}
            //if (!openFlg)
            //{
            //    Tokuchoyaro form = new Tokuchoyaro();
            //    form.UserInfos = this.UserInfos;
            //    form.Show();
            //    //this.Close();
            //}
            //this.Hide();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            this.ReSearch = true;
            //Madoguchi form = new Madoguchi();
            //form.UserInfos = this.UserInfos;
            //form.Show();
            //this.Close();
            Form f = null;
            Boolean openFlg = false;
            for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
            {
                f = System.Windows.Forms.Application.OpenForms[i];
                if (f.Text.IndexOf("窓口ミハル") >= 0 && f.Text.IndexOf("編集") <= -1)
                {
                    f.Show();
                    openFlg = true;
                    break;
                }
            }
            if (!openFlg)
            {
                Madoguchi form = new Madoguchi();
                form.UserInfos = this.UserInfos;
                form.Show();
                //this.Close();
            }
            this.Hide();
        }
        // 特命課長
        private void button3_Click(object sender, EventArgs e)
        {
            this.ReSearch = true;
            //Tokumei form = new Tokumei();
            //form.UserInfos = this.UserInfos;
            //form.Show();
            //this.Close();
            Form f = null;
            Boolean openFlg = false;
            for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
            {
                f = System.Windows.Forms.Application.OpenForms[i];
                if (f.Text.IndexOf("特命課長") >= 0 && f.Text.IndexOf("編集") <= -1)
                {
                    f.Show();
                    openFlg = true;
                    break;
                }
            }
            if (!openFlg)
            {
                Tokumei form = new Tokumei();
                form.UserInfos = this.UserInfos;
                form.Show();
                //this.Close();
            }
            this.Hide();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.ReSearch = true;
            //Jibun form = new Jibun();
            //form.UserInfos = this.UserInfos;
            //form.Show();
            //this.Close();
            Form f = null;
            Boolean openFlg = false;
            for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
            {
                f = System.Windows.Forms.Application.OpenForms[i];
                if (f.Text.IndexOf("自分大臣") >= 0 && f.Text.IndexOf("編集") <= -1)
                {
                    f.Show();
                    openFlg = true;
                    break;
                }
            }
            if (!openFlg)
            {
                Jibun form = new Jibun();
                form.UserInfos = this.UserInfos;
                form.Show();
                //this.Close();
            }
            this.Hide();
        }


        private void get_data(int mode = 0)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();

            if (src_Nendo.Text == "")
            {
                return;
            }
            string Nendo = src_Nendo.SelectedValue.ToString();

            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            try
            {
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    //担当者進捗
                    if (mode == 1)
                    {
                        //担当者進捗
                        cmd.CommandText = "SELECT " +
                            "MadoguchiJouhou.MadoguchiID " + //0:窓口ID
                            ",MadoguchiL1ChousaCD " + //1:調査CD
                            ", " +
                            "CASE " +
                            "WHEN MadoguchiHoukokuzumi = 1 THEN '8' " +
                            "WHEN MadoguchiHoukokuzumi != 1 THEN " +
                            "     CASE " +
                            //"         WHEN MadoguchiShinchokuJoukyou = 80 THEN '6' " +
                            //"         WHEN MadoguchiShinchokuJoukyou = 70 THEN '5' " +
                            //"         WHEN MadoguchiShinchokuJoukyou = 50 THEN '7' " +
                            //"         WHEN MadoguchiShinchokuJoukyou = 60 THEN '7' " + // 一次検済
                            "         WHEN MadoguchiL1ChousaShinchoku = 80 THEN '6' " +
                            "         WHEN MadoguchiL1ChousaShinchoku = 70 THEN '5' " +
                            "         WHEN MadoguchiL1ChousaShinchoku = 50 THEN '7' " +
                            "         WHEN MadoguchiL1ChousaShinchoku = 60 THEN '7' " + // 一次検済
                            "     ELSE " +
                            "         CASE " +
                            //"              WHEN MadoguchiShimekiribi < '" + DateTime.Today + "' THEN '1' " +
                            //"              WHEN MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(3) + "' THEN '2' " +
                            //"              WHEN MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(7) + "' THEN '3' " +
                            "              WHEN MadoguchiL1ChousaShimekiribi < '" + DateTime.Today + "' THEN '1' " +
                            "              WHEN MadoguchiL1ChousaShimekiribi <= '" + DateTime.Today.AddDays(3) + "' THEN '2' " +
                            "              WHEN MadoguchiL1ChousaShimekiribi <= '" + DateTime.Today.AddDays(7) + "' THEN '3' " +
                            "         ELSE '4' " +
                            "         END " +
                            "     END " +
                            "END " + //2:状況
                            ",MadoguchiUketsukeBangou + '-' + MadoguchiUketsukeBangouEdaban " + //3:特調番号
                            ",MadoguchiL1ChousaShinchoku " + //4:担当者進捗状況
                            ",MadoguchiL1ChousaShimekiribi " + //5:調査員締切日
                            ",MadoguchiHachuuKikanmei " + //5:発注者名・課名
                            ",CASE WHEN ISNULL(MadoguchiShukeiHyoFolder, '') = '' THEN 0 ELSE 1 END " + //6:集計表アイコン
                            ",MadoguchiShukeiHyoFolder " + //7:集計表パス
                            ",MadoguchiGyoumuMeishou " + //8:業務名称
                            ",MadoguchiL1Memo " + //9:メモ
                            "FROM MadoguchiJouhouMadoguchiL1Chou " +
                            "LEFT JOIN MadoguchiJouhou ON MadoguchiJouhou.MadoguchiID = MadoguchiJouhouMadoguchiL1Chou.MadoguchiID " +
                            "WHERE MadoguchiL1ChousaTantoushaCD = " + UserInfos[0] + " " +
                            //"AND (MadoguchiL1ChousaShinchoku <= 40 OR MadoguchiL1ChousaShinchoku = 80 )" +
                            //"AND (MadoguchiL1ChousaShinchoku <= 60 OR MadoguchiL1ChousaShinchoku = 80) " +
                            // 1200 調査担当者進捗状況で中止は非表示
                            "AND (MadoguchiL1ChousaShinchoku <= 60) " +
                            "AND MadoguchiTourokuNendo = " + Nendo + " " +
                            "AND MadoguchiDeleteFlag != 1 " +
                            "AND MadoguchiSystemRenban > 0 " +
                            // 1205 登録順で表示
                            //"ORDER BY MadoguchiJouhou.MadoguchiID ";
                            // 1205 締切日順らしい
                            "ORDER BY MadoguchiJouhouMadoguchiL1Chou.MadoguchiL1ChousaShimekiribi ";
                        Console.WriteLine(cmd.CommandText);
                        var sda = new SqlDataAdapter(cmd);
                        DT_UserShinchoku.Clear();
                        sda.Fill(DT_UserShinchoku);

                        //GRID初期化
                        c1FlexGrid1.Rows.Count = 1;
                    }
                    //窓口進捗状況
                    if (mode == 2)
                    {
                        //cmd.CommandText = "SELECT " +
                        //    "MadoguchiJouhou.MadoguchiID " + //0:窓口ID
                        //    "FROM MadoguchiJouhouMadoguchiL1Chou " +
                        //    "LEFT JOIN MadoguchiJouhou ON MadoguchiJouhou.MadoguchiID = MadoguchiJouhouMadoguchiL1Chou.MadoguchiID " +
                        //    "WHERE MadoguchiL1ChousaTantoushaCD = " + UserInfos[0] + " " +
                        //    "AND (MadoguchiL1ChousaShinchoku <= 40 OR MadoguchiL1ChousaShinchoku = 80 )" +
                        //    "AND MadoguchiTourokuNendo = " + Nendo + " ";
                        //Console.WriteLine(cmd.CommandText);

                        //担当者進捗
                        cmd.CommandText = "SELECT " +
                            "MadoguchiJouhou.MadoguchiID " + // 0:窓口ID
                            ", " +
                            "CASE " +
                            "WHEN MadoguchiHoukokuzumi = 1 THEN '8' " +
                            "WHEN MadoguchiHoukokuzumi != 1 THEN " +
                            "     CASE " +
                            "         WHEN MadoguchiShinchokuJoukyou = 80 THEN '6' " +
                            "         WHEN MadoguchiShinchokuJoukyou = 70 THEN '5' " +
                            "         WHEN MadoguchiShinchokuJoukyou = 50 THEN '7' " +
                            "         WHEN MadoguchiShinchokuJoukyou = 60 THEN '7' " + // 一次検済
                            "     ELSE " +
                            "         CASE " +
                            "              WHEN MadoguchiShimekiribi < '" + DateTime.Today + "' THEN '1' " +
                            "              WHEN MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(3) + "' THEN '2' " +
                            "              WHEN MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(7) + "' THEN '3' " +
                            "         ELSE '4' " +
                            "         END " +
                            "     END " +
                            "END " + // 1:状況
                            ",MadoguchiUketsukeBangou + '-' + MadoguchiUketsukeBangouEdaban " + // 2:特調番号
                            ",MadoguchiShimekiribi " + // 3:窓口締切日
                            ",MadoguchiHoukokuJisshibi " + // 4:報告実施日
                            ",CASE WHEN LEN(MadoguchiHachuuKikanmei) > 30 THEN SUBSTRING(MadoguchiHachuuKikanmei,1,30) ELSE MadoguchiHachuuKikanmei END AS MadoguchiHachuuKikanmei " + //5:発注者名・課名
                            "FROM MadoguchiJouhou " +
                            "WHERE MadoguchiHoukokuzumi = 0 " +
                            "AND MadoguchiTantoushaBushoCD = '" + UserInfos[2] + "' " +
                            "AND MadoguchiTantoushaCD = '" + UserInfos[0] + "' " + 
                            "AND (MadoguchiShinchokuJoukyou != 80 ) " +
                            "AND MadoguchiTourokuNendo = " + Nendo + " " +
                            "AND MadoguchiDeleteFlag != 1 " +
                            "AND MadoguchiSystemRenban > 0 " + 
                            "ORDER BY MadoguchiShimekiribi,MadoguchiShinchokuJoukyou,MadoguchiUketsukeBangou,MadoguchiUketsukeBangouEdaban";
                        Console.WriteLine(cmd.CommandText);

                        var sda = new SqlDataAdapter(cmd);
                        DT_MadoguchiShinshoku.Clear();
                        sda.Fill(DT_MadoguchiShinshoku);

                        //GRID初期化
                        c1FlexGrid2.Rows.Count = 1;
                    }
                    //担当者変更履歴
                    if (mode == 3)
                    {

                        cmd.CommandText = "SELECT " +
                            "th.MadoguchiID " +                  // 0:窓口ID
                             //",H_TOKUCHOBANGOU " +             // 1:特調番号
                             // 特調番号は、Historyからではなく、MadoguchiJouhou から出していたので修正
                            ",CASE WHEN mj.MadoguchiUketsukeBangouEdaban is null OR mj.MadoguchiUketsukeBangouEdaban = '' then mj.MadoguchiUketsukeBangou " + // 1:特調番号
                            " ELSE mj.MadoguchiUketsukeBangou + '-' + mj.MadoguchiUketsukeBangouEdaban END AS tokuchoNo " +
                            ",H_OPERATE_DT " +                // 2:操作日時
                            ",H_OPERATE_USER_MEI " +          // 3:操作者
                            ",H_OPERATE_NAIYO " +             // 4:変更内容
                            //",HistoryBeforeTantoubushoMei " + // 5:変更前担当者
                            //",HistoryAfterTantoubushoMei " +  // 6:変更後担当者
                            ",HistoryBeforeTantoushaMei " + // 5:変更前担当者
                            ",HistoryAfterTantoushaMei " +  // 6:変更後担当者
                            "FROM T_HISTORY th " +
                            "INNER JOIN MadoguchiJouhou mj ON mj.MadoguchiID = th.MadoguchiID " +
                            "WHERE th.MadoguchiID IS NOT NULL " +
                            "AND HistoryAfterTantoubushoCD = '" + UserInfos[2] + "' " +
                            "AND ISNULL(HistoryBeforeTantoushaCD, '') <> '' " +
                            "AND ISNULL(HistoryAfterTantoushaCD, '') <> '' " +
                            "AND (HistoryBeforeTantoushaCD = '" + UserInfos[0] + "' OR HistoryAfterTantoushaCD = '" + UserInfos[0] + "') " +
                            "AND H_DATE_KEY > '" + DateTime.Today.AddDays(-10) + "' " +
                            "AND mj.MadoguchiDeleteFlag != 1 " +
                            "AND MadoguchiSystemRenban > 0 " +
                            "AND mj.MadoguchiTourokuNendo = " + Nendo + " " + // 売上年度で絞る
                            "ORDER BY H_OPERATE_DT DESC ";

                        var sda = new SqlDataAdapter(cmd);
                        DT_TantoushaChange.Clear();
                        sda.Fill(DT_TantoushaChange);

                        //GRID初期化
                        c1FlexGrid3.Rows.Count = 1;
                    }
                    if (mode == 4)
                    {
                        if (src_ShuFuku.Text == "")
                        {
                            return;
                        }
                        int CHOUSA_CHRNGE = 0;
                        if (GlobalMethod.GetCommonValue1("CHOUSA_CHRNGE") == "1")
                        {
                            CHOUSA_CHRNGE = 1;
                        }

                        ////担当者空白リスト
                        //cmd.CommandText = "SELECT " +
                        //    "MadoguchiJouhouMadoguchiL1Chou.MadoguchiID " + //0:窓口ID
                        //    //",MadoguchiL1ChousaCD " + //1:調査CD
                        //    ", " +
                        //    "CASE " +
                        //    "WHEN MadoguchiHoukokuzumi = 1 THEN '8' " +
                        //    "WHEN MadoguchiHoukokuzumi != 1 THEN " +
                        //    "     CASE " +
                        //    "         WHEN MadoguchiShinchokuJoukyou = 80 THEN '6' " +
                        //    "         WHEN MadoguchiShinchokuJoukyou = 70 THEN '5' " +
                        //    "         WHEN MadoguchiShinchokuJoukyou = 50 THEN '7' " +
                        //    "     ELSE " +
                        //    "         CASE " +
                        //    "              WHEN MadoguchiShimekiribi < '" + DateTime.Today + "' THEN '1' " +
                        //    "              WHEN MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(3) + "' THEN '2' " +
                        //    "              WHEN MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(7) + "' THEN '3' " +
                        //    "         ELSE '4' " +
                        //    "         END " +
                        //    "     END " +
                        //    "END " + //2:状況
                        //    ",MadoguchiUketsukeBangou + '-' + MadoguchiUketsukeBangouEdaban " + //3:特調番号
                        //    ",MadoguchiL1ChousaShinchoku " + //4:担当者進捗状況
                        //    ",MadoguchiTourokubi " + //5:登録日
                        //    ",MadoguchiShimekiribi " + //5:窓口締切日
                        //    ",MadoguchiL1ChousaBushoCD " + //5:窓口部所
                        //    ",MadoguchiHachuuKikanmei " + //5:発注者名・課名
                        //    "FROM MadoguchiJouhou " +
                        //    "LEFT JOIN  MadoguchiJouhouMadoguchiL1Chou ON MadoguchiJouhou.MadoguchiID = MadoguchiJouhouMadoguchiL1Chou.MadoguchiID " +
                        //    "INNER JOIN ChousaHinmoku ON MadoguchiJouhou.MadoguchiID = ChousaHinmoku.MadoguchiID ";
                        //if (CHOUSA_CHRNGE == 1)
                        //{
                        //    cmd.CommandText += "WHERE (MadoguchiShinchokuJoukyou <= 50 OR MadoguchiShinchokuJoukyou = 80) " +
                        //                        "AND (ChousaShinchokuJoukyou <= 50 OR ChousaShinchokuJoukyou = 80) ";
                        //}
                        //else
                        //{
                        //    cmd.CommandText += "WHERE (MadoguchiShinchokuJoukyou < 50 OR MadoguchiShinchokuJoukyou = 80) " +
                        //                        "AND (ChousaShinchokuJoukyou < 50 OR ChousaShinchokuJoukyou = 80) ";
                        //}
                        //cmd.CommandText += "AND MadoguchiHoukokuzumi = 0 " +
                        //    "AND MadoguchiTourokuNendo = " + Nendo + " AND ISNULL(ChousaTanka, '') <> '' ";

                        ////担当者主副コンボ　0：主担当
                        //if (src_ShuFuku.SelectedValue.ToString() == "0")
                        //{
                        //    cmd.CommandText += "AND ISNULL(HinmokuChousainCD, '') = '' AND (HinmokuRyakuBushoCD = '" + UserInfos[2] + "' OR MadoguchiTantoushaBushoCD = '" + UserInfos[2] + "') ";

                        //}
                        //// 1:副担当
                        //else if (src_ShuFuku.SelectedValue.ToString() == "1")
                        //{
                        //    cmd.CommandText += "AND ((MadoguchiTantoushaBushoCD = '" + UserInfos[2] + "' AND (" +
                        //                        "(HinmokuFukuChousainCD1 IS NULL AND ISNULL(HinmokuRyakuBushoFuku1CD, '') <> '') OR " +
                        //                        "(HinmokuFukuChousainCD2 IS NULL AND ISNULL(HinmokuRyakuBushoFuku2CD, '') <> '') )) " +
                        //                        " OR (HinmokuFukuChousainCD1 IS NULL AND HinmokuRyakuBushoFuku1CD = '" + UserInfos[2] + "') " +
                        //                        " OR (HinmokuFukuChousainCD2 IS NULL AND HinmokuRyakuBushoFuku2CD = '" + UserInfos[2] + "')) ";
                        //}
                        //else if (src_ShuFuku.SelectedValue.ToString() == "4")
                        //{
                        //    cmd.CommandText += "AND ((MadoguchiTantoushaBushoCD = '" + UserInfos[2] + "' AND (" +
                        //                        "(HinmokuFukuChousainCD1 IS NULL AND ISNULL(HinmokuRyakuBushoFuku1CD, '') <> '') OR " +
                        //                        "(HinmokuFukuChousainCD2 IS NULL AND ISNULL(HinmokuRyakuBushoFuku2CD, '') <> '') )) " +
                        //                        " OR ( (ISNULL(HinmokuChousainCD, '') = '' AND HinmokuRyakuBushoCD = '" + UserInfos[2] + "') " +
                        //                        " OR (HinmokuFukuChousainCD1 IS NULL AND HinmokuRyakuBushoFuku1CD = '" + UserInfos[2] + "') " +
                        //                        " OR (HinmokuFukuChousainCD2 IS NULL AND HinmokuRyakuBushoFuku2CD = '" + UserInfos[2] + "'))) ";
                        //}

                        //    cmd.CommandText += "ORDER BY ChousaHinmokuShimekiribi DESC, MadoguchiJouhouMadoguchiL1Chou.MadoguchiID";

                        //担当者空白リスト
                        cmd.CommandText = "SELECT " +
                            "MadoguchiJouhou.MadoguchiID " + //0:窓口ID
                                                             //",MadoguchiL1ChousaCD " + //1:調査CD
                            ", " +
                            "CASE " +
                            "WHEN MadoguchiHoukokuzumi = 1 THEN '8' " +
                            "WHEN MadoguchiHoukokuzumi != 1 THEN " +
                            "     CASE " +
                            "         WHEN MadoguchiShinchokuJoukyou = 80 THEN '6' " +
                            "         WHEN MadoguchiShinchokuJoukyou = 70 THEN '5' " +
                            "         WHEN MadoguchiShinchokuJoukyou = 50 THEN '7' " +
                            "         WHEN MadoguchiShinchokuJoukyou = 60 THEN '7' " + // 一次検済
                            "     ELSE " +
                            "         CASE " +
                            "              WHEN MadoguchiShimekiribi < '" + DateTime.Today + "' THEN '1' " +
                            "              WHEN MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(3) + "' THEN '2' " +
                            "              WHEN MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(7) + "' THEN '3' " +
                            "         ELSE '4' " +
                            "         END " +
                            "     END " +
                            "END " + //1:状況
                            ",MadoguchiUketsukeBangou + '-' + MadoguchiUketsukeBangouEdaban " + //2:特調番号
                            ",ChousaShinchokuJoukyou " + //3:担当者進捗状況
                            ",MadoguchiTourokubi " + //4:登録日
                            ",MadoguchiShimekiribi " + //5:窓口締切日
                            ",MadoguchiTantoushaBushoCD " + //6:窓口部所
                            ",MadoguchiHachuuKikanmei " + //7:発注者名・課名
                            "FROM MadoguchiJouhou " +
                            "INNER JOIN ChousaHinmoku ON MadoguchiJouhou.MadoguchiID = ChousaHinmoku.MadoguchiID " +
                            "WHERE MadoguchiHoukokuzumi = 0 " +
                            "AND MadoguchiTourokuNendo = " + src_Nendo.SelectedValue + " ";
                        if (CHOUSA_CHRNGE == 1)
                        {
                            cmd.CommandText += "AND (MadoguchiShinchokuJoukyou <= 50 OR MadoguchiShinchokuJoukyou = 80) " +
                                                "AND (ChousaShinchokuJoukyou <= 50 OR ChousaShinchokuJoukyou = 80) ";
                        }
                        else
                        {
                            cmd.CommandText += "AND (MadoguchiShinchokuJoukyou < 50 OR MadoguchiShinchokuJoukyou = 80) " +
                                                "AND (ChousaShinchokuJoukyou < 50 OR ChousaShinchokuJoukyou = 80) ";
                        }
                        //cmd.CommandText += "AND MadoguchiHoukokuzumi = 0 " +
                        //cmd.CommandText += "AND MadoguchiTourokuNendo = " + Nendo + " AND ISNULL(ChousaTanka, '') <> '' ";
                        cmd.CommandText += "AND ISNULL(ChousaTanka, '') <> '' ";

                        //担当者主副コンボ　0：主担当
                        if (src_ShuFuku.SelectedValue.ToString() == "0")
                        {
                            cmd.CommandText += "AND ISNULL(HinmokuChousainCD, '') = '' AND (HinmokuRyakuBushoCD = '" + UserInfos[2] + "' OR MadoguchiTantoushaBushoCD = '" + UserInfos[2] + "') ";

                        }
                        // 1:副担当
                        else if (src_ShuFuku.SelectedValue.ToString() == "1")
                        {
                            cmd.CommandText += "AND ((MadoguchiTantoushaBushoCD = '" + UserInfos[2] + "' AND (" +
                                                "(HinmokuFukuChousainCD1 IS NULL AND ISNULL(HinmokuRyakuBushoFuku1CD, '') <> '') OR " +
                                                "(HinmokuFukuChousainCD2 IS NULL AND ISNULL(HinmokuRyakuBushoFuku2CD, '') <> '') )) " +
                                                " OR (HinmokuFukuChousainCD1 IS NULL AND HinmokuRyakuBushoFuku1CD = '" + UserInfos[2] + "') " +
                                                " OR (HinmokuFukuChousainCD2 IS NULL AND HinmokuRyakuBushoFuku2CD = '" + UserInfos[2] + "')) ";
                        }
                        else if (src_ShuFuku.SelectedValue.ToString() == "4")
                        {
                            cmd.CommandText += "AND ((MadoguchiTantoushaBushoCD = '" + UserInfos[2] + "' AND (" +
                                                "(HinmokuFukuChousainCD1 IS NULL AND ISNULL(HinmokuRyakuBushoFuku1CD, '') <> '') OR " +
                                                "(HinmokuFukuChousainCD2 IS NULL AND ISNULL(HinmokuRyakuBushoFuku2CD, '') <> '') )) " +
                                                " OR ( (ISNULL(HinmokuChousainCD, '') = '' AND HinmokuRyakuBushoCD = '" + UserInfos[2] + "') " +
                                                " OR (HinmokuFukuChousainCD1 IS NULL AND HinmokuRyakuBushoFuku1CD = '" + UserInfos[2] + "') " +
                                                " OR (HinmokuFukuChousainCD2 IS NULL AND HinmokuRyakuBushoFuku2CD = '" + UserInfos[2] + "'))) ";
                        }
                        // 808 担当者空白リストから、部所のみで中止の場合、表示しない。
                        // →中止を抽出しない
                        //cmd.CommandText += "AND (ChousaChuushi != 1) ";
                        //cmd.CommandText += "AND ((HinmokuRyakuBushoCD != null OR HinmokuRyakuBushoFuku1CD != null OR HinmokuRyakuBushoFuku2CD != null) ";
                        //cmd.CommandText += "AND HinmokuChousainCD = null AND HinmokuFukuChousainCD1 = null AND HinmokuFukuChousainCD2 = null AND ChousaChuushi != 1) ";
                        cmd.CommandText += "AND (ChousaChuushi != 1) "; // 調査品目で中止となっているものを除外
                        cmd.CommandText += "AND (MadoguchiJiishiKubun != 3) "; // 実施区分を中止を除外
                        cmd.CommandText += "AND (MadoguchiDeleteFlag != 1) ";
                        cmd.CommandText += "AND (MadoguchiSystemRenban > 0) ";

                        cmd.CommandText += "ORDER BY ChousaHinmokuShimekiribi DESC, MadoguchiJouhou.MadoguchiID";
                        Console.WriteLine(cmd.CommandText);
                        var sda = new SqlDataAdapter(cmd);
                        DT_BlankShinchoku.Clear();
                        sda.Fill(DT_BlankShinchoku);

                        //GRID初期化
                        c1FlexGrid4.Rows.Count = 1;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            set_data(mode);

            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void set_data(int mode)
        {
            if (mode == 1 && DT_UserShinchoku != null)
            {
                //描画停止
                c1FlexGrid1.BeginUpdate();

                for (int i = 0; i < DT_UserShinchoku.Rows.Count; i++)
                {
                    c1FlexGrid1.Rows.Add();
                    for (int k = 0; k < c1FlexGrid1.Cols.Count; k++)
                    {
                        c1FlexGrid1[i + 1, k] = DT_UserShinchoku.Rows[i][k];
                    }
                    if (DT_UserShinchoku.Rows[i][6].ToString() == "1" && !Directory.Exists(DT_UserShinchoku.Rows[i][7].ToString()))
                    {
                        c1FlexGrid1[i + 1, 6] = 0;
                    }
                    c1FlexGrid1.Rows[i + 1].Height = 28;
                }

                //描画再開
                c1FlexGrid1.EndUpdate();
            }
            if (mode == 2 && DT_UserShinchoku != null)
            {
                //描画停止
                c1FlexGrid2.BeginUpdate();

                for (int i = 0; i < DT_MadoguchiShinshoku.Rows.Count; i++)
                {
                    c1FlexGrid2.Rows.Add();
                    for (int k = 0; k < c1FlexGrid2.Cols.Count; k++)
                    {
                        c1FlexGrid2[i + 1, k] = DT_MadoguchiShinshoku.Rows[i][k];
                    }
                    c1FlexGrid2.Rows[i + 1].Height = 28;
                }

                //描画再開
                c1FlexGrid2.EndUpdate();
            }
            if (mode == 3 && DT_TantoushaChange != null)
            {
                //描画停止
                c1FlexGrid3.BeginUpdate();

                for (int i = 0; i < DT_TantoushaChange.Rows.Count; i++)
                {
                    c1FlexGrid3.Rows.Add();
                    for (int k = 0; k < c1FlexGrid3.Cols.Count; k++)
                    {
                        c1FlexGrid3[i + 1, k] = DT_TantoushaChange.Rows[i][k];
                    }
                    c1FlexGrid3.Rows[i + 1].Height = 28;
                }

                //描画再開
                c1FlexGrid3.EndUpdate();
            }
            if (mode == 4 && DT_BlankShinchoku != null)
            {
                //MadoguchiID重複データを削除
                DataView dv = new DataView(DT_BlankShinchoku);
                DT_BlankShinchoku = dv.ToTable(true);

                //描画停止
                c1FlexGrid4.BeginUpdate();

                for (int i = 0; i < DT_BlankShinchoku.Rows.Count; i++)
                {
                    c1FlexGrid4.Rows.Add();
                    for (int k = 0; k < c1FlexGrid4.Cols.Count; k++)
                    {
                        c1FlexGrid4[i + 1, k] = DT_BlankShinchoku.Rows[i][k];
                    }
                    c1FlexGrid4.Rows[i + 1].Height = 28;
                }

                //描画再開
                c1FlexGrid4.EndUpdate();
            }
        }

        private void set_combo()
        {

            //年度
            string discript = "NendoSeireki";
            string value = "NendoID";
            string table = "Mst_Nendo";
            string where = "";
            //コンボボックスデータ取得
            DataTable tmpdt = GlobalMethod.getData(discript, value, table, where);
            src_Nendo.DisplayMember = "Discript";
            src_Nendo.ValueMember = "Value";
            src_Nendo.DataSource = tmpdt;

            //検索条件初期化
            //売上年度　受託課所支部
            /*
            discript = "NendoSeireki ";
            value = "NendoID ";
            table = "Mst_Nendo ";
            where = "Nendo_Sdate <= GETDATE() AND Nendo_EDate >= GETDATE()";
            //コンボボックスデータ取得
            DataTable dt = GlobalMethod.getData(discript, value, table, where);
            if (dt != null)
            {
                src_Nendo.SelectedValue = dt.Rows[0][0].ToString();
            }
            else
            {
                src_Nendo.SelectedValue = System.DateTime.Now.Year;
            }
            */
            src_Nendo.SelectedValue = GlobalMethod.GetTodayNendo();

            //空白リスト主副
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "主担当");
            tmpdt.Rows.Add(1, "副担当");
            tmpdt.Rows.Add(4, "主＋副担当");
            src_ShuFuku.DisplayMember = "Discript";
            src_ShuFuku.ValueMember = "Value";
            src_ShuFuku.DataSource = tmpdt;


            // 進捗アイコン
            Hashtable imgMap = new Hashtable();
            imgMap.Add("8", Image.FromFile("Resource/Image/shin_ao.png"));     // 報告済み
            imgMap.Add("5", Image.FromFile("Resource/Image/greenT1.png"));     // 二次検証済み、または中止（二次検証済み）
            //imgMap.Add("6", Image.FromFile("Resource/Image/greenT1.png"));     // 二次検証済み、または中止（中止）
            imgMap.Add("6", Image.FromFile("Resource/Image/shin_ao.png"));     // 中止
            imgMap.Add("7", Image.FromFile("Resource/Image/shin_midori.png")); // 担当者済み
            imgMap.Add("1", Image.FromFile("Resource/Image/shin_dokuro.png")); // 締切日経過
            imgMap.Add("2", Image.FromFile("Resource/Image/shin_aka.png"));    // 締切日が3日以内、かつ2次検証が完了していない
            imgMap.Add("3", Image.FromFile("Resource/Image/shin_kiiro.png"));  // 締切日が1週間以内、かつ2次検証が完了していない
            imgMap.Add("4", Image.FromFile("Resource/Image/blank.png"));      // 上記のいずれにも該当しない
            c1FlexGrid1.Cols[2].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[2].ImageMap = imgMap;
            c1FlexGrid1.Cols[2].ImageAndText = false;
            c1FlexGrid2.Cols[1].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid2.Cols[1].ImageMap = imgMap;
            c1FlexGrid2.Cols[1].ImageAndText = false;
            c1FlexGrid4.Cols[1].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid4.Cols[1].ImageMap = imgMap;
            c1FlexGrid4.Cols[1].ImageAndText = false;

            //フォルダの画像切り替え
            imgMap = new Hashtable();
            imgMap.Add("0", Image.FromFile("Resource/Image/folder_gray_s.png"));
            imgMap.Add("1", Image.FromFile("Resource/Image/folder_yellow_s.png"));
            c1FlexGrid1.Cols[7].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[7].ImageMap = imgMap;
            c1FlexGrid1.Cols[7].ImageAndText = false;

            // 担当者状況
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "");
            tmpdt.Rows.Add(10, "依頼");
            tmpdt.Rows.Add(20, "調査開始");
            tmpdt.Rows.Add(30, "見積中");
            tmpdt.Rows.Add(40, "集計中");
            tmpdt.Rows.Add(50, "担当者済");
            tmpdt.Rows.Add(60, "一次検済");
            tmpdt.Rows.Add(70, "二次検済");
            tmpdt.Rows.Add(80, "中止");
            SortedList sl = GlobalMethod.Get_SortedList(tmpdt);

            c1FlexGrid1.Cols[4].DataMap = sl;
            c1FlexGrid4.Cols[3].DataMap = sl;

            // Gridの受託部所の為の設定
            discript = "Mst_Busho.ShibuMei";
            value = "Mst_Busho.GyoumuBushoCD ";
            table = "Mst_Busho";
            where = "";
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            c1FlexGrid4.Cols[6].DataMap = sl; // 受託部所



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

        private void src_ShuFuku_SelectedIndexChanged(object sender, EventArgs e)
        {
            get_data(4);
        }

        //調査担当者進捗状況Grid押下時
        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid1.HitTest(new Point(e.X, e.Y));
            //特徴番号列　自分大臣詳細に遷移
            if (hti.Row > 0 && hti.Column == 3)
            {
                this.ReSearch = true;
                Jibun_Input form = new Jibun_Input();
                form.MadoguchiID = c1FlexGrid1[hti.Row, 0].ToString();
                form.ChousaCD = c1FlexGrid1[hti.Row, 1].ToString();
                form.UserInfos = UserInfos;
                form.Show(this);
                this.Hide();

                //this.ReSearch = true;

                //Form f = null;
                //Boolean openFlg = false;
                //for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
                //{
                //    f = System.Windows.Forms.Application.OpenForms[i];
                //    if (f.Text.IndexOf("自分大臣") >= 0 && f.Text.IndexOf("編集") <= -1)
                //    {
                //        Jibun_Input form = (Jibun_Input)f;
                //        form.MadoguchiID = c1FlexGrid1[hti.Row, 0].ToString();
                //        form.ChousaCD = c1FlexGrid1[hti.Row, 1].ToString();
                //        form.UserInfos = UserInfos;
                //        form.Show(this);
                //        openFlg = true;
                //        break;
                //    }
                //}
                //if (!openFlg)
                //{
                //    Jibun_Input form = new Jibun_Input();
                //    form.MadoguchiID = c1FlexGrid1[hti.Row, 0].ToString();
                //    form.ChousaCD = c1FlexGrid1[hti.Row, 1].ToString();
                //    form.UserInfos = UserInfos;
                //    form.Show(this);
                //}
                //this.Hide();
            }
            //集計表フォルダ表示
            if (hti.Row > 0 && hti.Column == 7 && c1FlexGrid1[hti.Row, 7].ToString().Equals("1"))
            {
                System.Diagnostics.Process.Start("EXPLORER.EXE", GlobalMethod.GetPathValid(c1FlexGrid1[hti.Row, 8].ToString()));
            }
        }

        private void src_Nendo_SelectedIndexChanged(object sender, EventArgs e)
        {
            get_data(1);
            get_data(2);
            get_data(3);
            get_data(4);
        }

        //調査担当者進捗状況　編集後
        private void c1FlexGrid1_AfterEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            if (e.Col == 10)
            {
                string MadoguchiID = c1FlexGrid1.Rows[e.Row][0].ToString();
                string ChousaCD = c1FlexGrid1.Rows[e.Row][1].ToString();
                string Memo = c1FlexGrid1.Rows[e.Row][10].ToString();

                GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "調査担当者メモ更新 MadoguchiID:" + MadoguchiID + " ChousaCD:" + ChousaCD, "Update_Tokuchoyaro_Memo", MadoguchiID);
                string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                SqlConnection sqlconn = new SqlConnection(connStr);
                sqlconn.Open();
                SqlTransaction transaction = sqlconn.BeginTransaction();
                var cmd = sqlconn.CreateCommand();
                cmd.Transaction = transaction;
                try
                {
                    cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET  " +
                        "MadoguchiL1Memo = N'" + Memo + "' " +
                        ",MadoguchiL1UpdateDate = SYSDATETIME() " +
                        ",MadoguchiL1UpdateUser = N'" + UserInfos[0] + "' " +
                        ",MadoguchiL1UpdateProgram = 'Update_Tokuchoyaro_Memo' " +
                         " WHERE MadoguchiID = '" + MadoguchiID + "' AND MadoguchiL1ChousaCD = '" + ChousaCD + "' ";

                    Console.WriteLine(cmd.CommandText);
                    cmd.ExecuteNonQuery();
                    transaction.Commit();

                }
                catch
                {
                    transaction.Rollback();
                }
                finally
                {
                    sqlconn.Close();
                }
            }
        }

        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }


        // 窓口業務別進捗状況Grid
        private void c1FlexGrid2_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid2.HitTest(new Point(e.X, e.Y));
            // 特調番号列　窓口ミハル詳細に遷移
            if (hti.Row > 0 && hti.Column == 2)
            {
                this.ReSearch = true;
                Madoguchi_Input form = new Madoguchi_Input();
                form.MadoguchiID = c1FlexGrid2[hti.Row, 0].ToString();
                form.UserInfos = UserInfos;
                form.Show(this);
                this.Hide();

                //this.ReSearch = true;

                //Form f = null;
                //Boolean openFlg = false;
                //for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
                //{
                //    f = System.Windows.Forms.Application.OpenForms[i];
                //    if (f.Text.IndexOf("窓口ミハル") >= 0 && f.Text.IndexOf("編集") <= -1)
                //    {
                //        Madoguchi_Input form = (Madoguchi_Input)f;
                //        form.MadoguchiID = c1FlexGrid2[hti.Row, 0].ToString();
                //        form.UserInfos = UserInfos;
                //        form.Show(this);
                //        openFlg = true;
                //        break;
                //    }
                //}
                //if (!openFlg)
                //{
                //    Madoguchi_Input form = new Madoguchi_Input();
                //    form.MadoguchiID = c1FlexGrid1[hti.Row, 0].ToString();
                //    form.UserInfos = UserInfos;
                //    form.Show(this);
                //}
                //this.Hide();
            }
        }

        // 調査品目担当者変更履歴（過去10日以内）Grid
        private void c1FlexGrid3_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid3.HitTest(new Point(e.X, e.Y));
            // 特調番号列　窓口ミハル詳細に遷移
            if (hti.Row > 0 && hti.Column == 1)
            {
                this.ReSearch = true;
                Jibun_Input form = new Jibun_Input();
                form.MadoguchiID = c1FlexGrid3[hti.Row, 0].ToString();
                form.UserInfos = UserInfos;
                form.Show(this);
                this.Hide();
                //this.ReSearch = true;

                //Form f = null;
                //Boolean openFlg = false;
                //for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
                //{
                //    f = System.Windows.Forms.Application.OpenForms[i];
                //    if (f.Text.IndexOf("自分大臣") >= 0 && f.Text.IndexOf("編集") <= -1)
                //    {
                //        Jibun_Input form = (Jibun_Input)f;
                //        form.MadoguchiID = c1FlexGrid3[hti.Row, 0].ToString();
                //        form.UserInfos = UserInfos;
                //        form.Show();
                //        openFlg = true;
                //        break;
                //    }
                //}
                //if (!openFlg)
                //{
                //    Jibun_Input form = new Jibun_Input();
                //    form.MadoguchiID = c1FlexGrid3[hti.Row, 0].ToString();
                //    form.UserInfos = UserInfos;
                //    form.Show(this);
                //}
                //this.Hide();

            }
        }

        //担当者空白リストGrid押下時
        private void c1FlexGrid4_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid4.HitTest(new Point(e.X, e.Y));
            //特徴番号列　窓口ミハル詳細に遷移
            if (hti.Row > 0 && hti.Column == 2)
            {
                this.ReSearch = true;
                Madoguchi_Input form = new Madoguchi_Input();
                form.MadoguchiID = c1FlexGrid4[hti.Row, 0].ToString();
                form.UserInfos = UserInfos;
                form.Show(this);
                this.Hide();


                //this.ReSearch = true;

                //Form f = null;
                //Boolean openFlg = false;
                //for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
                //{
                //    f = System.Windows.Forms.Application.OpenForms[i];
                //    if (f.Text.IndexOf("窓口ミハル") >= 0 && f.Text.IndexOf("編集") <= -1)
                //    {
                //        Madoguchi_Input form = (Madoguchi_Input)f;
                //        form.MadoguchiID = c1FlexGrid4[hti.Row, 0].ToString();
                //        form.UserInfos = UserInfos;
                //        form.Show(this);
                //        openFlg = true;
                //        break;
                //    }
                //}
                //if (!openFlg)
                //{
                //    Madoguchi_Input form = new Madoguchi_Input();
                //    form.MadoguchiID = c1FlexGrid4[hti.Row, 0].ToString();
                //    form.UserInfos = UserInfos;
                //    form.Show(this);
                //}
                //this.Hide();
            }
        }

        // 検索ボタン
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            get_data(1);
            get_data(2);
            get_data(3);
            get_data(4);
        }

        private void Tokuchoyaro_Activated(object sender, EventArgs e)
        {
            if (ReSearch)
            {
                get_data(1);
                get_data(2);
                get_data(3);
                get_data(4);
                ReSearch = false;
            }
        }
    }
}
