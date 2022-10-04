using C1.C1Excel;
using C1.Win.C1FlexGrid;
using C1.Win.C1Input;
using C1.Win.C1Input.GrapeCity.Editors;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;

namespace TokuchoBugyoK2
{
    public partial class tanka : Form
    {
        string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
        GlobalMethod GlobalMethod = new GlobalMethod();
        public String[] UserInfos;
        private string AnkenID = "";
        DataTable TankaRank_Data = new DataTable();
        DataTable Jimusho_Data = new DataTable();
        DataTable Tantousya_Data = new DataTable();
        DataTable Busho_Data = new DataTable();
        DataTable CommonMaster_Data = new DataTable();
        DataTable CommonMaster_Data_RankRow = new DataTable();
        //DataTable TankaRank_Data_Work = new DataTable();
        //DataTable TankaRank_Data_Save = new DataTable();
        int intTankaRankRow = 0;
        private int TankaKeiyakuID = 0;
        private int TankaKeiyakuID_Main = 0;
        private int TankaKeiyakuID_Copy = 0;
        private string JutakuBangou = "";

        private DateTime NullDate;

        public tanka()
        {
            InitializeComponent();

            //マウスホイールの制御を追加
            this.HoukokuSentaku.MouseWheel += item_MouseWheel;
            this.PrintList.MouseWheel += item_MouseWheel;
            this.ShuKeiHouhou.MouseWheel += item_MouseWheel;
        }

        private void Grid_AfterResizeRow(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            string Name = ((C1FlexGrid)sender).Name;
            Resize_Grid(Name);
        }

        private void tanka_Load(object sender, EventArgs e)
        {
            //不具合No1355（1123）
            lblVersion.Text = GlobalMethod.GetCommonValue1("APL_VERSION");
            if (GlobalMethod.GetCommonValue1("BOOT_MODE") == "1")
            {
                lblBootMode.Text = GlobalMethod.GetCommonValue2("BOOT_MODE");
            }
            // 昇順降順アイコン設定
            TankaRankuGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            TankaRankuGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            KoujiJimusyoGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            KoujiJimusyoGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            TantoushaGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            TantoushaGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            label1.Text = UserInfos[3] + "：" + UserInfos[1];

            set_combo();

        }

        private void set_combo()
        {
            //単価ランクの設定
            System.Data.DataTable tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));

            tmpdt.Rows.Add(1, "合計");
            tmpdt.Rows.Add(2, "最大");
            SortedList sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //該当グリッドのセルにセット
            TankaRankuGrid.Cols[3].DataMap = sl;
            DataRow dr;
            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }

            ShuKeiHouhou.DataSource = tmpdt;
            ShuKeiHouhou.DisplayMember = "Discript";
            ShuKeiHouhou.ValueMember = "Value";

            //帳票リスト
            string BushoCD = "";
            String discript = "ShuukeiMei";
            String value = "ShuukeiMei";
            String table = "Mst_Busho";
            String where = "GyoumuBushoCD = '" + UserInfos[2] + "' AND ShuukeiMei = '1.本部' ";
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);

            if (combodt != null && combodt.Rows.Count > 0)
            {
                BushoCD = combodt.Rows[0][0].ToString();
            }
            else
            {
                BushoCD = UserInfos[2].ToString();
            }

            
            discript = "PrintName";
            value = "PrintListID";
            table = "Mst_PrintList";
            where = "MENU_ID = 208 AND PrintDelFlg <> '1' AND (BushoKanriboBushoCD = '" + BushoCD + "' or BushoKanriboBushoCD is null) ORDER BY PrintListNarabijun ";
            combodt = GlobalMethod.getData(discript, value, table, where);
            //if (combodt != null)
            //{
            //    dr = combodt.NewRow();
            //    combodt.Rows.InsertAt(dr, 0);
            //}
            PrintList.DataSource = combodt;
            PrintList.DisplayMember = "Discript";
            PrintList.ValueMember = "Value";

            //報告選択
            System.Data.DataTable tmpdt2 = new System.Data.DataTable();
            tmpdt2.Columns.Add("Value", typeof(int));
            tmpdt2.Columns.Add("Discript", typeof(string));
            tmpdt2.Clear();
            tmpdt2.Rows.Add(1, "報告日");
            tmpdt2.Rows.Add(2, "報告実施日");
            if (tmpdt2 != null)
            {
                dr = tmpdt2.NewRow();
                tmpdt2.Rows.InsertAt(dr, 0);
            }

            HoukokuSentaku.DataSource = tmpdt2;
            HoukokuSentaku.DisplayMember = "Discript";
            HoukokuSentaku.ValueMember = "Value";

        }

        private void get_data(string AnkenJouhouID)
        {
            try
            {
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();

                    //単価ランクの取得
                    //cmd.CommandText = "SELECT " +
                    //  "TankaRankHinmoku,TankaRankKakaku,TankaRankShubetsu " +
                    //  "FROM TankaKeiyakuRank " +
                    //  "WHERE TankaKeiyakuID = '" + AnkenJouhouID + "' ORDER BY TankaRankHinmoku";
                    //cmd.CommandText = "SELECT"
                    //                + " TankaRankHinmoku,TankaRankKakaku,TankaRankShubetsu,TankaKeiyakuRank.TankaKeiyakuID"
                    //                + " FROM TankaKeiyakuRank"
                    //                + " LEFT JOIN TankaKeiyaku ON TankaKeiyakuRank.TankaKeiyakuID = TankaKeiyaku.TankaKeiyakuID"
                    //                + " WHERE AnkenJouhouID = '" + AnkenJouhouID + "' ORDER BY TankaRankHinmoku"
                    //                ;
                    //cmd.CommandText = "SELECT"
                    //                + " tkr.TankaRankHinmoku,tkr.TankaRankKakaku,tkr.TankaRankShubetsu,tkr.TankaKeiyakuID"
                    //                + " FROM TankaKeiyakuRank tkr"
                    //                + " LEFT JOIN TankaKeiyaku tk ON tk.TankaKeiyakuID = tkr.TankaKeiyakuID"
                    //                + " LEFT JOIN AnkenJouhou aj ON aj.AnkenJutakuBangou = tk.TankakeiyakuJutakuBangou"
                    //                + " WHERE aj.AnkenJouhouID = '" + AnkenJouhouID + "' ORDER BY TankaRankHinmoku"
                    //                ;

                    cmd.CommandText = "SELECT"
                                    + " tkr.TankaRankHinmoku,tkr.TankaRankKakaku,tkr.TankaRankShubetsu,tkr.TankaKeiyakuID"
                                    + " FROM TankaKeiyakuRank tkr"
                                    ;
                    if (TankaKeiyakuID != 0)
                    {
                        cmd.CommandText += " WHERE tkr.TankaKeiyakuID = '" + TankaKeiyakuID + "'"
                                         ;
                    }
                    else
                    {
                        cmd.CommandText += " LEFT JOIN TankaKeiyaku tk ON tk.TankaKeiyakuID = tkr.TankaKeiyakuID"
                                         + " LEFT JOIN AnkenJouhou aj ON aj.AnkenJutakuBangou = tk.TankakeiyakuJutakuBangou"
                                         + " WHERE aj.AnkenJouhouID = '" + AnkenJouhouID + "'"
                                         ;
                    }
                    cmd.CommandText += " ORDER BY tkr.TankaRankHinmoku";

                    var sda = new SqlDataAdapter(cmd);
                    TankaRank_Data.Clear();
                    sda.Fill(TankaRank_Data);

                    //TankaKeiyakuID = 0;
                    if (TankaRank_Data.Rows.Count > 0 && TankaRank_Data.Rows[0][3] != null)
                    {
                        TankaKeiyakuID = int.Parse(TankaRank_Data.Rows[0][3].ToString());
                    }

                    // 共通マスタから単価ランクの初期行数（TANKAKEIYAKU_RANK_ROW）の取得
                    cmd.CommandText = "SELECT"
                                    + " CommonValue1, CommonValue1Type, CommonValue2, CommonValue2Type"
                                    + " FROM M_CommonMaster"
                                    + " WHERE CommonMasterKye = 'TANKAKEIYAKU_RANK_ROW'"
                                    + " ORDER BY CommonMasterID"
                                    ;

                    sda = new SqlDataAdapter(cmd);
                    CommonMaster_Data_RankRow.Clear();
                    sda.Fill(CommonMaster_Data_RankRow);

                    if (CommonMaster_Data_RankRow.Rows.Count != 0)
                    {
                        intTankaRankRow = int.Parse(CommonMaster_Data_RankRow.Rows[0][0].ToString());
                    }

                    //データがない場合、共通マスタから取得した値で初期の行を用意する
                    if (TankaRank_Data.Rows.Count == 0)
                    {
                        ////退避データがあった場合、そちらから設定する
                        //if (TankaRank_Data_Save.Rows.Count > 0)
                        //{
                        //    TankaRank_Data = TankaRank_Data_Save;
                        //}
                        //else
                        //{
                            ////共通マスタから単価ランクの初期行数（TANKAKEIYAKU_RANK_ROW）の取得
                            //cmd.CommandText = "SELECT " +
                            //  "CommonValue1,CommonValue1Type,CommonValue2,CommonValue2Type " +
                            //  "FROM M_CommonMaster " +
                            //  "WHERE CommonMasterKye = 'TANKAKEIYAKU_RANK_ROW' ORDER BY CommonMasterID";

                            //sda = new SqlDataAdapter(cmd);
                            //CommonMaster_Data.Clear();
                            //sda.Fill(CommonMaster_Data);

                            //if (CommonMaster_Data.Rows.Count != 0)
                            //{
                            //    intTankaRankRow = int.Parse(CommonMaster_Data.Rows[0][0].ToString());
                            //}

                            //共通マスタから単価ランク名称の初期値（TANKAKEIYAKU_RANK_INITIAL）の取得
                            cmd.CommandText = "SELECT " +
                              "CommonValue1,CommonValue1Type,CommonValue2,CommonValue2Type " +
                              "FROM M_CommonMaster " +
                              "WHERE CommonMasterKye = 'TANKAKEIYAKU_RANK_INITIAL' ORDER BY CommonMasterID";

                            sda = new SqlDataAdapter(cmd);
                            CommonMaster_Data.Clear();
                            sda.Fill(CommonMaster_Data);
                        //}
                    }

                    //事務所の取得
                    //if (AnkenID != AnkenJouhouID)
                    //{
                    //    cmd.CommandText = "SELECT " +
                    //      "'',KoujijimushoMei,KoujijimushoYomi,KoujijimushoUketsukeNo,KoujijimushoTantouYakushoku " +
                    //      "FROM Mst_Koujijimusho " +
                    //      //"WHERE TankaKeiyakuID = '" + AnkenJouhouID + "' ";
                    //      "WHERE TankaKeiyakuID = '" + TankaKeiyakuID + "' ";
                    //}
                    //else
                    //{
                        cmd.CommandText = "SELECT " +
                          "KoujijimushoID,KoujijimushoMei,KoujijimushoYomi,KoujijimushoUketsukeNo,KoujijimushoTantouYakushoku " +
                          "FROM Mst_Koujijimusho " +
                          //"WHERE TankaKeiyakuID = '" + AnkenJouhouID + "' ";
                          "WHERE TankaKeiyakuID = '" + TankaKeiyakuID + "' " +
                          " ORDER BY KoujijimushoMei";
                    //}

                    sda = new SqlDataAdapter(cmd);
                    Jimusho_Data.Clear();
                    sda.Fill(Jimusho_Data);

                    //担当者の取得
                    cmd.CommandText = "SELECT " +
                      //"KoujijimushoMei,KoujiTantoushaBusho,KoujiTantoushaYakushoku,KoujiTantoushaMei,KoujiTantoushaTEL,KoujiTantoushaFAX,KoujiTantoushaMail " +
                      "KoujijimushoMei,KoujiTantoushaBusho,KoujiTantoushaYakushoku,KoujiTantoushaMei,KoujiTantoushaTEL,KoujiTantoushaFAX,KoujiTantoushaMail,Mst_KoujijimushoTantousha.KoujijimushoID " +
                      "FROM Mst_KoujijimushoTantousha " +
                      "LEFT JOIN Mst_Koujijimusho ON Mst_KoujijimushoTantousha.KoujijimushoID = Mst_Koujijimusho.KoujijimushoID " +
                      //"WHERE TankaKeiyakuID = '" + AnkenJouhouID + "' ";
                      "WHERE TankaKeiyakuID = '" + TankaKeiyakuID + "' " +
                      " ORDER BY KoujijimushoMei,KoujiTantoushaBusho,KoujiTantoushaMei";

                    sda = new SqlDataAdapter(cmd);
                    Tantousya_Data.Clear();
                    sda.Fill(Tantousya_Data);

                    // 工事事務所の行番号に一次的に工事事務所IDをセットしているので行番号に書き換える
                    for (int i = 0; i < Tantousya_Data.Rows.Count; i++)
                    {
                        string KoujijimushoID = Tantousya_Data.Rows[i][7].ToString();
                        for (int k = 0; k < Jimusho_Data.Rows.Count; k++)
                        {
                            if (Jimusho_Data.Rows[k][0].ToString() == KoujijimushoID)
                            {
                                int JimushoGridNo = k + 1;
                                Tantousya_Data.Rows[i][7] = JimushoGridNo.ToString();
                            }
                        }
                    }

                    // コピーの場合、工事事務所IDをクリアする
                    if (AnkenID != AnkenJouhouID)
                    {
                        for (int i = 0; i < Jimusho_Data.Rows.Count; i++)
                        {
                            Jimusho_Data.Rows[i][0] = 0;
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            set_data();
        }

        private void set_data()
        {
            //単価ランクの設定
            TankaRankuGrid.Rows.Count = 1;
            for (int i = 0; i < TankaRank_Data.Rows.Count; i++)
            {
                TankaRankuGrid.Rows.Add();
                // VIPS 20220414 コンポーネント最新化にあたり修正
                //for (int k = 0; k < TankaRank_Data.Columns.Count; k++)
                for (int k = 0; k < TankaRank_Data.Columns.Count - 1; k++)
                {
                    TankaRankuGrid.Rows[i + 1][k + 1] = TankaRank_Data.Rows[i][k].ToString();
                }
            }

            ////データがない場合、共通マスタから取得した値で初期の行を用意する
            //if (TankaRank_Data.Rows.Count == 0)
            //{
            //    //共通マスタから初期値が取得できなかった場合は、空行を追加する
            //    if (CommonMaster_Data.Rows.Count == 0)
            //    {
            //        for (int i = 0; i < intTankaRankRow; i++)
            //        {
            //            TankaRankuGrid.Rows.Add();
            //            TankaRankuGrid.Rows[i + 1][3] = "1";                                        //集計方法
            //        }
            //    }
            //    else
            //    {
            //        for (int i = 0; i < CommonMaster_Data.Rows.Count; i++)
            //        {
            //            TankaRankuGrid.Rows.Add();
            //            TankaRankuGrid.Rows[i + 1][1] = CommonMaster_Data.Rows[i][0].ToString();    //単価ランク名称
            //            TankaRankuGrid.Rows[i + 1][3] = "1";                                        //集計方法
            //        }
            //    }
            //}
            // 単価ランクが取得できていない、かつ共通マスタから初期値が取得できた場合は共通マスタから値をセットする
            if (CommonMaster_Data.Rows.Count != 0 && TankaRank_Data.Rows.Count == 0)
            {
                for (int i = 0; i < CommonMaster_Data.Rows.Count; i++)
                {
                    TankaRankuGrid.Rows.Add();
                    TankaRankuGrid.Rows[i + 1][1] = CommonMaster_Data.Rows[i][0].ToString();    //単価ランク名称
                    TankaRankuGrid.Rows[i + 1][3] = "1";                                        //集計方法
                }
            }

            // 共通マスタで設定されている初期行数よりGridに設定されている行数が少ない場合、不足分を設定する
            if (TankaRankuGrid.Rows.Count - 1 < intTankaRankRow)
            {
                int num = 0;
                if (TankaRankuGrid.Rows.Count != 1)
                {
                    num = TankaRankuGrid.Rows.Count - 1;
                }
                for (int i = num; i < intTankaRankRow; i++)
                {
                    TankaRankuGrid.Rows.Add();
                    TankaRankuGrid.Rows[i + 1][3] = "1";                                            //集計方法
                }
            }
            Resize_Grid("TankaRankuGrid");

            //事務所の設定
            KoujiJimusyoGrid.Rows.Count = 1;
            for (int i = 0; i < Jimusho_Data.Rows.Count; i++)
            {
                KoujiJimusyoGrid.Rows.Add();
                for (int k = 0; k < Jimusho_Data.Columns.Count; k++)
                {
                    KoujiJimusyoGrid.Rows[i + 1][k + 1] = Jimusho_Data.Rows[i][k];
                }
            }
            Resize_Grid("KoujiJimusyoGrid");

            //担当者の設定
            TantoushaGrid.Rows.Count = 1;
            for (int i = 0; i < Tantousya_Data.Rows.Count; i++)
            {
                TantoushaGrid.Rows.Add();
                for (int k = 0; k < Tantousya_Data.Columns.Count; k++)
                {
                    TantoushaGrid.Rows[i + 1][k + 1] = Tantousya_Data.Rows[i][k];
                }
            }
            Resize_Grid("TantoushaGrid");

        }

        //private void save_data() 
        //{
        //    string TankaRankHinmoku = "";
        //    string TankaRankKakaku = "";
        //    string TankaRankShubetsu = "";

        //    TankaRank_Data_Save.Clear();
        //    //グリッドに値がセットされていたら退避する
        //    if (TankaRankuGrid.Rows.Count > 1)
        //    {
        //        TankaRank_Data_Work.Clear();
        //        TankaRank_Data_Work.Columns.Clear();

        //        TankaRank_Data_Work.Columns.Add("TankaRankHinmoku");
        //        TankaRank_Data_Work.Columns.Add("TankaRankKakaku");
        //        TankaRank_Data_Work.Columns.Add("TankaRankShubetsu");

        //        for (int i = 1; i < TankaRankuGrid.Rows.Count; i++)
        //        {
        //            TankaRankHinmoku = "";
        //            TankaRankKakaku = "";
        //            TankaRankShubetsu = "";

        //            if (TankaRankuGrid.Rows[i][1] != null)
        //            {
        //                TankaRankHinmoku = TankaRankuGrid.Rows[i][1].ToString();
        //            }
        //            if (TankaRankuGrid.Rows[i][2] != null)
        //            {
        //                TankaRankKakaku = TankaRankuGrid.Rows[i][2].ToString();
        //            }
        //            if (TankaRankuGrid.Rows[i][3] != null)
        //            {
        //                TankaRankShubetsu = TankaRankuGrid.Rows[i][3].ToString();
        //            }
        //            TankaRank_Data_Work.Rows.Add(TankaRankHinmoku, TankaRankKakaku, TankaRankShubetsu);
        //        }
        //        //直接セットすると値が参照渡しになっていて元のTankaRank_Dataがクリアされた際に一緒に消えたので、コピーでセットする
        //        TankaRank_Data_Save = TankaRank_Data_Work.Copy();
        //    }
        //}

        public void Resize_Grid(string name)
        {
            Control[] cs;
            cs = this.Controls.Find(name, true);
            if (cs.Length > 0)
            {
                var fx = (C1FlexGrid)cs[0];
                int h = 0;
                for (int i = 0; i < fx.Rows.Count; i++)
                {
                    if (fx.Rows[i].Visible == true)
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
                }
                fx.Height = 4 + h;

                int w = 0;
                for (int i = 0; i < fx.Cols.Count; i++)
                {
                    if (fx.Cols[i].Visible == true)
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
                }
                if (fx.Width < 4 + w)
                {
                    fx.Height += 18;
                }
            }
        }

        private Boolean Chk_data()
        {
            Boolean ErrorFlag = false;

            // 価格設定時に単価ランク名称が未設定の場合、エラー
            for (int i = 1; i < TankaRankuGrid.Rows.Count; i++)
            {
                if (TankaRankuGrid.Rows[i][2] != null && decimal.Parse(TankaRankuGrid.Rows[i][2].ToString()) > 0)
                {
                    if (TankaRankuGrid.Rows[i][1] == null || TankaRankuGrid.Rows[i][1].ToString() == "")
                    {
                        //"必須入力項目が入力されていません。"('E20805')
                        ErrorFlag = true;
                        set_error(GlobalMethod.GetMessage("E20805", ""));
                        break;
                    }
                }
            }
            //単価ランクと工事事務所
            //E20806
            //単価ランクエラーチェック
            var TankaRankList = new List<string>();
            for (int i = 1; i < TankaRankuGrid.Rows.Count; i++)
            {
                if (TankaRankuGrid.Rows[i][1] != null && TankaRankuGrid.Rows[i][1].ToString() != "")
                {
                    //単価ランク重複
                    if (TankaRankList.IndexOf(TankaRankuGrid.Rows[i][1].ToString()) == -1)
                    {
                        TankaRankList.Add(TankaRankuGrid.Rows[i][1].ToString());
                    }
                    else
                    {
                        //"単価ランク名称が重複しています。"(E20803)
                        set_error(GlobalMethod.GetMessage("E20803", ""));
                        ErrorFlag = true;
                        break;
                    }

                    //単価ランク文字数
                    if (TankaRankuGrid.Rows[i][1].ToString().Length > 15)
                    {
                        //単価ランク名は15文字以内で入力して下さい。(E20807)
                        set_error(GlobalMethod.GetMessage("E20807", ""));
                        ErrorFlag = true;
                        break;
                    }


                }
            }

            //調査品目で使用されている単価ランクがいない
            if (TankaKeiyakuID_Main != 0)
            {
                //DataTable ChousaHinmoku = GlobalMethod.getData("ChousaHoukokuRank", "DISTINCT ChousaHoukokuRank", "ChousaHinmoku LEFT JOIN MadoguchiJouhou ON ChousaHinmoku.MadoguchiID = MadoguchiJouhou.MadoguchiID", "AnkenJouhouID = " + AnkenID + " AND ChousaHoukokuRank <> '' ");
                DataTable ChousaHinmoku = GlobalMethod.getData("ChousaHoukokuRank", "DISTINCT ChousaHoukokuRank", "ChousaHinmoku LEFT JOIN TanpinNyuuryoku ON ChousaHinmoku.MadoguchiID = TanpinNyuuryoku.MadoguchiID", "TanpinGyoumuCD = " + TankaKeiyakuID_Main + " AND ChousaHoukokuRank <> '' ");
                for (int i = 0; i < ChousaHinmoku.Rows.Count; i++)
                {
                    if (TankaRankList.IndexOf(ChousaHinmoku.Rows[i][1].ToString()) == -1)
                    {
                        //調査品目に設定されている単価ランク名が変更、あるいは、削除されました。調査品目に設定されたランク名は変更、削除が出来ません。(E20808)"　変更・削除されたランク名 = "
                        set_error(GlobalMethod.GetMessage("E20808", "変更・削除されたランク名 = " + ChousaHinmoku.Rows[i][1].ToString()));
                        ErrorFlag = true;
                    }
                }
            }
            if (ErrorFlag)
            {
                //"工事事務所一覧の更新に失敗しました。")E20804
                set_error(GlobalMethod.GetMessage("E20801", ""));
                return false;
            }

            return true;
        }

        private void c1FlexGrid1_BeforeMouseDown(object sender, BeforeMouseDownEventArgs e)
        {
            var hti = ((C1FlexGrid)sender).HitTest(new System.Drawing.Point(e.X, e.Y));
            if (hti.Row > 0 && hti.Column == 0)
            {
                if (GlobalMethod.outputMessage("I10002", "", 1) == DialogResult.OK)
                {
                    TankaRankuGrid.Rows.Remove(hti.Row);
                    Resize_Grid("TankaRankuGrid");
                }
            }
        }

        private void c1FlexGrid2_BeforeMouseDown(object sender, BeforeMouseDownEventArgs e)
        {
            set_error("", 0);
            var hti = ((C1FlexGrid)sender).HitTest(new System.Drawing.Point(e.X, e.Y));
            if (hti.Row > 0 && hti.Column == 0)
            {
                if (GlobalMethod.outputMessage("I10002", "", 1) == DialogResult.OK)
                {
                    int tantousya = 0;
                    for (int i = 1; i < TantoushaGrid.Rows.Count; i++)
                    {
                        if (TantoushaGrid.Rows[i][1].ToString() == KoujiJimusyoGrid.Rows[hti.Row][2].ToString())
                        {
                            tantousya++;
                        }
                    }
                    if (tantousya > 0)
                    {
                        //担当者を先に削除しないとエラー
                        set_error("事務所を削除する前に、担当者を削除してください。");
                    }
                    else
                    {
                        KoujiJimusyoGrid.Rows.Remove(hti.Row);
                        Resize_Grid("KoujiJimusyoGrid");
                    }
                }
            }
        }

        private void c1FlexGrid3_BeforeMouseDown(object sender, BeforeMouseDownEventArgs e)
        {
            var hti = ((C1FlexGrid)sender).HitTest(new System.Drawing.Point(e.X, e.Y));
            if (hti.Row > 0 && hti.Column == 0)
            {
                if (GlobalMethod.outputMessage("I10002", "", 1) == DialogResult.OK)
                {
                    TantoushaGrid.Rows.Remove(hti.Row);
                    Resize_Grid("TantoushaGrid");
                }
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
                string txt = e.Index > -1 ? ((ComboBox)sender).Items[e.Index].ToString() : ((ComboBox)sender).Text;
                e.Graphics.DrawString(txt, e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            }
            e.DrawFocusRectangle();
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

        private void button_Update_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            //単価契約更新処理
            //データチェック処理
            if (!Chk_data())
            {
                return;
            }


            //更新処理
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();

                SqlTransaction transaction = conn.BeginTransaction();
                cmd.Transaction = transaction;

                try
                {
                    //// 受託番号選択時に単価契約が無かった場合
                    //if ("0".Equals(TankaKeiyakuID_Main) || "".Equals(TankaKeiyakuID_Main))
                    //{
                    //    DataTable TankaKeiyaku_Dt = new DataTable();
                    //    //cmd.CommandText = "SELECT TankaKeiyakuID FROM TankaKeiyaku"
                    //                    //+ " WHERE AnkenJouhouID = " + AnkenID
                    //    cmd.CommandText = "SELECT tk.TankaKeiyakuID FROM TankaKeiyaku tk"
                    //                    + " LEFT JOIN AnkenJouhou aj ON aj.AnkenJutakuBangou = tk.TankakeiyakuJutakuBangou"
                    //                    + " WHERE aj.AnkenJouhouID = '" + AnkenID + "'"
                    //                    + " ORDER BY tk.TankaKeiyakuID DESC" + // 単価契約が存在しないので、最大値
                    //                    ;

                    //    var sda = new SqlDataAdapter(cmd);
                    //    sda.Fill(TankaKeiyaku_Dt);

                    //    TankaKeiyakuID = 0;
                    //    if (TankaKeiyaku_Dt.Rows.Count > 0 && TankaKeiyaku_Dt.Rows[0][0] != null)
                    //    {
                    //        //TankaKeiyakuID = int.Parse(TankaKeiyaku_Dt.Rows[0][0].ToString());
                    //        TankaKeiyakuID_Main = int.Parse(TankaKeiyaku_Dt.Rows[0][0].ToString());
                    //    }
                    //}
                    TankaKeiyakuID = TankaKeiyakuID_Main; // 受託番号選択で取得したIDをセット

                    //単価契約テーブルに登録/更新
                    //if (GlobalMethod.Check_Table(AnkenID, "AnkenJouhouID", "TankaKeiyaku", ""))
                    if (TankaKeiyakuID != 0)
                    {
                        cmd.CommandText = "UPDATE TankaKeiyaku SET " +
                                            " TankakeiyakuUpdateDate = GETDATE() " +
                                            ",TankakeiyakuUpdateUser = '" + UserInfos[0] + "' " +
                                            ",TankakeiyakuUpdateProgram = 'UpdateTanka' " +
                                            //" WHERE AnkenJouhouID = " + AnkenID;
                                            " WHERE TankaKeiyakuID = " + TankaKeiyakuID;
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                    }
                    else
                    {

                        TankaKeiyakuID = GlobalMethod.getSaiban("TankaKeiyakuID");

                        cmd.CommandText = "INSERT INTO TankaKeiyaku ( " +
                                            "TankaKeiyakuID " +
                                            ",AnkenJouhouID " +
                                            ",TankakeiyakuJutakuBangou " +
                                            ",TankakeiyakuCreateDate " +
                                            ",TankakeiyakuCreateUser " +
                                            ",TankakeiyakuCreateProgram " +
                                            ",TankakeiyakuUpdateDate " +
                                            ",TankakeiyakuUpdateUser " +
                                            ",TankakeiyakuUpdateProgram " +
                                            ",TankakeiyakuDeleteFlag " +
                                            " ) VALUES ( " +
                                            //AnkenID +
                                            TankaKeiyakuID +
                                            "," + AnkenID +
                                            ",N'" + GlobalMethod.ChangeSqlText(Header_JutakuBangou.Text, 0) + "' " +
                                            ",  GETDATE() " +
                                            ", '" + UserInfos[0] + "'" +
                                            ", 'InsertTanka'" +
                                            ",  GETDATE() " +
                                            ", '" + UserInfos[0] + "'" +
                                            ", 'InsertTanka'" +
                                            ",'0' " +
                                            " )";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                    }


                    //単価ランクテーブルに登録
                    //cmd.CommandText = "DELETE FROM TankaKeiyakuRank	WHERE TankaKeiyakuID = " + AnkenID;
                    cmd.CommandText = "DELETE FROM TankaKeiyakuRank	WHERE TankaKeiyakuID = " + TankaKeiyakuID;
                    cmd.ExecuteNonQuery();

                    for (int i = 1; i < TankaRankuGrid.Rows.Count; i++)
                    {
                        if (TankaRankuGrid.Rows[i][1] != null && TankaRankuGrid.Rows[i][1].ToString() != "")
                        {
                            cmd.CommandText = "INSERT INTO TankaKeiyakuRank ( " +
                                                "TankaKeiyakuID " +
                                                ",TankaRankID " +
                                                ",TankaRankHinmoku " +
                                                ",TankaRankKakaku " +
                                                ",TankaRankShubetsu " +
                                                ",TankaRankCreateDate " +
                                                ",TankaRankCreateUser " +
                                                ",TankaRankCreateProgram " +
                                                ",TankaRankUpdateDate " +
                                                ",TankaRankUpdateUser " +
                                                ",TankaRankUpdateProgram " +
                                                ",TankaRankDeleteFlag " +
                                                " ) VALUES ( " +
                                                //AnkenID +
                                                TankaKeiyakuID +
                                                "," + i +
                                                ",N'" + GlobalMethod.ChangeSqlText(TankaRankuGrid.Rows[i][1].ToString(), 0) + "' " +
                                                ",N'" + TankaRankuGrid.Rows[i][2] + "' " +
                                                ",N'" + GlobalMethod.ChangeSqlText(TankaRankuGrid.Rows[i][3].ToString(), 0) + "' " +
                                                ",  GETDATE() " +
                                                ", N'" + UserInfos[0] + "'" +
                                                ", 'InsertTanka'" +
                                                ",  GETDATE() " +
                                                ", N'" + UserInfos[0] + "'" +
                                                ", 'InsertTanka'" +
                                                ",'0' " +
                                                " )";
                            Console.WriteLine(cmd.CommandText);
                            cmd.ExecuteNonQuery();

                            // 新規登録時のTankaKeiykauIDを保持しておく
                            TankaKeiyakuID_Main = TankaKeiyakuID;
                        }
                    }

                    //工事事務所テーブルに登録
                    /*
                    for (int i = 1; i < c1FlexGrid2.Rows.Count; i++)
                    {
                        if (c1FlexGrid2.Rows[i][1] != null && c1FlexGrid2.Rows[i][1].ToString() != "")
                        {
                            cmd.CommandText = "DELETE FROM Mst_KoujijimushoTantousha WHERE KoujijimushoID = " + c1FlexGrid2.Rows[i][1];
                            Console.WriteLine(cmd.CommandText);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    */
                    //cmd.CommandText = "DELETE T FROM Mst_KoujijimushoTantousha AS T LEFT JOIN Mst_Koujijimusho K ON K.KoujijimushoID = T.KoujijimushoID WHERE TankaKeiyakuID = " + AnkenID;
                    cmd.CommandText = "DELETE T FROM Mst_KoujijimushoTantousha AS T LEFT JOIN Mst_Koujijimusho K ON K.KoujijimushoID = T.KoujijimushoID WHERE TankaKeiyakuID = " + TankaKeiyakuID;
                    Console.WriteLine(cmd.CommandText);
                    cmd.ExecuteNonQuery();
                    //cmd.CommandText = "DELETE FROM Mst_Koujijimusho	WHERE TankaKeiyakuID = " + AnkenID;
                    cmd.CommandText = "DELETE FROM Mst_Koujijimusho	WHERE TankaKeiyakuID = " + TankaKeiyakuID;
                    Console.WriteLine(cmd.CommandText);
                    cmd.ExecuteNonQuery();

                    string JimusyoID = "";
                    for (int i = 1; i < KoujiJimusyoGrid.Rows.Count; i++)
                    {
                        JimusyoID = "";
                        //事務所IDがGridに存在するか
                        if (KoujiJimusyoGrid.Rows[i][1] != null && KoujiJimusyoGrid.Rows[i][1].ToString() != "0")
                        {
                            JimusyoID = KoujiJimusyoGrid.Rows[i][1].ToString();
                            /*
                            cmd.CommandText = "UPDATE Mst_Koujijimusho SET " +
                                                " KoujijimushoMei = '" + c1FlexGrid2.Rows[i][2] + "' " +
                                                ",KoujijimushoYomi = '" + c1FlexGrid2.Rows[i][3] + "' " +
                                                ",KoujijimushoUketsukeNo = '" + c1FlexGrid2.Rows[i][4] + "' " +
                                                ",KoujijimushoTantouYakushoku = '" + c1FlexGrid2.Rows[i][5] + "' " +
                                                ",KoujijimushoUpdateDate = GETDATE() " +
                                                ",KoujijimushoUpdateUser = '" + UserInfos[0] + "' " +
                                                ",KoujijimushoUpdateProgram = 'UpdateTanka' " +
                                                ",KoujijimushoDeleteFlag = 0 " +
                                                " WHERE KoujijimushoID = " + c1FlexGrid2.Rows[i][1].ToString();
                                                */
                        }
                        else
                        {
                            //工事事務所マスタに存在する名称か
                            /*
                            DataTable dt = GlobalMethod.getData("KoujijimushoMei", "KoujijimushoID", "Mst_Koujijimusho", "KoujijimushoMei = '" + c1FlexGrid2.Rows[i][2] + "'");
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                JimusyoID = dt.Rows[0][0].ToString();
                            }
                            else
                            {
                                JimusyoID = GlobalMethod.getSaiban("KoujijimushoID").ToString();
                            }
                            */
                            JimusyoID = GlobalMethod.getSaiban("KoujijimushoID").ToString();
                            KoujiJimusyoGrid.Rows[i][1] = JimusyoID;
                            //JimusyoID = GlobalMethod.getSaiban("HachuumotoKikanID").ToString(); 現行の採番処理はここ見てる

                        }
                        cmd.CommandText = "INSERT INTO Mst_Koujijimusho ( " +
                                            "KoujijimushoID " +
                                            ",KoujijimushoMei " +
                                            ",KoujijimushoYomi " +
                                            ",KoujijimushoUketsukeNo " +
                                            ",KoujijimushoTantouYakushoku " +
                                            ",TankaKeiyakuID " +
                                            ",KoujijimushoCreateDate " +
                                            ",KoujijimushoCreateUser " +
                                            ",KoujijimushoCreateProgram " +
                                            ",KoujijimushoUpdateDate " +
                                            ",KoujijimushoUpdateUser " +
                                            ",KoujijimushoUpdateProgram " +
                                            ",KoujijimushoDeleteFlag " +
                                            " ) VALUES ( " +
                                            "'" + JimusyoID + "' " +
                                            ",N'" + GlobalMethod.ChangeSqlText(KoujiJimusyoGrid.Rows[i][2].ToString(), 0) + "' " +
                                            ",N'" + GlobalMethod.ChangeSqlText(KoujiJimusyoGrid.Rows[i][3].ToString(), 0) + "' " +
                                            ",N'" + GlobalMethod.ChangeSqlText(KoujiJimusyoGrid.Rows[i][4].ToString(), 0) + "' " +
                                            ",N'" + GlobalMethod.ChangeSqlText(KoujiJimusyoGrid.Rows[i][5].ToString(), 0) + "' " +
                                            //"," + AnkenID +
                                            "," + TankaKeiyakuID +
                                            ",  GETDATE() " +
                                            ", N'" + UserInfos[0] + "'" +
                                            ", 'InsertTanka'" +
                                            ",  GETDATE() " +
                                            ", N'" + UserInfos[0] + "'" +
                                            ", 'InsertTanka'" +
                                            ",'0' " +
                                            " )";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                    }

                    //担当者テーブルに登録

                    //string JimusyoMei = "";
                    int JimushoGridNo = 0;
                    JimusyoID = "";
                    for (int i = 1; i < TantoushaGrid.Rows.Count; i++)
                    {
                        int.TryParse(TantoushaGrid.Rows[i][8].ToString(), out JimushoGridNo);
                        JimusyoID = KoujiJimusyoGrid.Rows[JimushoGridNo][1].ToString();
                        //if (JimusyoMei != TantoushaGrid.Rows[i][1].ToString())
                        //{
                        //    JimusyoMei = TantoushaGrid.Rows[i][1].ToString();
                        //    for (int k = 1; k < KoujiJimusyoGrid.Rows.Count; k++)
                        //    {
                        //        if (JimusyoMei == KoujiJimusyoGrid.Rows[k][2].ToString())
                        //        {
                        //            JimusyoID = KoujiJimusyoGrid.Rows[k][1].ToString();
                        //            break;
                        //        }
                        //    }
                            if (JimusyoID == "")
                            {
                                //事務所IDの取得失敗エラー
                                set_error(GlobalMethod.GetMessage("E20804", ""));
                                GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "工事事務所IDの取得に失敗しました", "InsertTankaKeiyaku", "");
                                transaction.Rollback();
                                conn.Close();
                                return;
                            }
                        //}
                        cmd.CommandText = "INSERT INTO Mst_KoujijimushoTantousha ( " +
                                            "KoujijimushoID " +
                                            ",KoujiTantoushaID " +
                                            ",KoujiTantoushaBusho " +
                                            ",KoujiTantoushaYakushoku " +
                                            ",KoujiTantoushaMei " +
                                            ",KoujiTantoushaTEL " +
                                            ",KoujiTantoushaFAX " +
                                            ",KoujiTantoushaMail " +
                                            ",KoujiTantoushaCreateDate " +
                                            ",KoujiTantoushaCreateUser " +
                                            ",KoujiTantoushaCreateProgram " +
                                            ",KoujiTantoushaUpdateDate " +
                                            ",KoujiTantoushaUpdateUser " +
                                            ",KoujiTantoushaUpdateProgram " +
                                            ",KoujiTantoushaDeleteFlag " +
                                            " ) VALUES ( " +
                                            "'" + JimusyoID + "' " +
                                            ",'" + i + "' " +
                                            ",N'" + GlobalMethod.ChangeSqlText(TantoushaGrid.Rows[i][2].ToString(), 0) + "' " +
                                            ",N'" + GlobalMethod.ChangeSqlText(TantoushaGrid.Rows[i][3].ToString(), 0) + "' " +
                                            ",N'" + GlobalMethod.ChangeSqlText(TantoushaGrid.Rows[i][4].ToString(), 0) + "' " +
                                            ",N'" + GlobalMethod.ChangeSqlText(TantoushaGrid.Rows[i][5].ToString(), 0) + "' " +
                                            ",N'" + GlobalMethod.ChangeSqlText(TantoushaGrid.Rows[i][6].ToString(), 0) + "' " +
                                            ",N'" + GlobalMethod.ChangeSqlText(TantoushaGrid.Rows[i][7].ToString(), 0) + "' " +
                                            ",  GETDATE() " +
                                            ", N'" + UserInfos[0] + "'" +
                                            ", 'InsertTanka'" +
                                            ",  GETDATE() " +
                                            ", N'" + UserInfos[0] + "'" +
                                            ", 'InsertTanka'" +
                                            ",'0' " +
                                            " )";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                    }

                    // 受託番号を元に窓口情報を参照し、業務CDが未設定のものがあれば更新する。
                    if (TankaKeiyakuID != 0)
                    {
                        cmd.CommandText = "UPDATE TanpinNyuuryoku "
                                        + " SET TanpinGyoumuCD = TankaKeiyakuID"
                                        + " FROM TanpinNyuuryoku "
                                        + " LEFT JOIN ("
                                        + "  SELECT MadoguchiID, (MadoguchiJutakuBangou +'-' + MadoguchiJutakuBangouEdaban) AS JutakuBangou"
                                        + "  FROM MadoguchiJouhou) AS MadoguchiJouhou ON TanpinNyuuryoku.MadoguchiID = MadoguchiJouhou.MadoguchiID"
                                        + " LEFT JOIN TankaKeiyaku ON MadoguchiJouhou.JutakuBangou = TankakeiyakuJutakuBangou"
                                        + " WHERE TanpinGyoumuCD = 0 AND TankaKeiyakuID = " + TankaKeiyakuID;

                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                    }

                    transaction.Commit();
                    //更新が正常に終了しました。I20801
                    set_error(GlobalMethod.GetMessage("I20803", ""));
                    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "単価契約を更新しました", "InsertTankaKeiyaku", "");
                }
                catch (Exception)
                {
                    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "単価契約を失敗しました", "InsertTankaKeiyaku", "");
                    set_error(GlobalMethod.GetMessage("E20801", ""));
                    transaction.Rollback();
                    //更新出来ませんでしたE20801
                    throw;
                }
                finally
                {
                    conn.Close();
                }

            }
            get_data(AnkenID);

        }

        private void button_Select_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            //TankaKeiyakuID = 0;
            Popup_Anken form = new Popup_Anken();
            form.mode = "tanka";
            /*
            String discript = "NendoSeireki";
            String value = "NendoID ";
            String table = "Mst_Nendo";
            String where = "Nendo_Sdate <= GETDATE() AND Nendo_EDate >= GETDATE()";
            DataTable dt = GlobalMethod.getData(discript, value, table, where);
            if (dt != null)
            {
                form.nendo = dt.Rows[0][0].ToString();
            }
            else
            {
                form.nendo = DateTime.Today.Year.ToString();
            }
            */
            form.nendo = GlobalMethod.GetTodayNendo();
            form.hachuushaKaMei = "";
            form.gyoumuMei = "";
            form.gyoumuBushoCD = UserInfos[2];

            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                AnkenID = form.ReturnValue[0];//案件ID
                Header_JutakuBangou.Text = form.ReturnValue[2];//受託番号
                JutakuBangou = form.ReturnValue[2].ToString().Substring(0, 9);
                Header_JutakuBusho.Text = form.ReturnValue[15];//受託部所
                Header_HachuushaKamei.Text = form.ReturnValue[14];//発注者名・課名
                Header_GyoumuMei.Text = form.ReturnValue[4];//業務名称
                int.TryParse(form.ReturnValue[16], out TankaKeiyakuID);//単価契約ID
                TankaKeiyakuID_Main = TankaKeiyakuID;

                if (AnkenID != "")
                {
                    //save_data();
                    get_data(AnkenID);
                    button_Update.Enabled = true;
                    button_Update.BackColor = Color.FromArgb(42, 78, 122);
                    //button3.Enabled = true;
                    //button3.BackColor = Color.FromArgb(42, 78, 122);
                    button_Copy.Enabled = true;
                    button_Copy.BackColor = Color.FromArgb(42, 78, 122);
                    button_Ranku.Enabled = true;
                    button_Ranku.BackColor = Color.FromArgb(42, 78, 122);
                    button_KoujiJimusyo.Enabled = true;
                    button_KoujiJimusyo.BackColor = Color.FromArgb(42, 78, 122);
                    button_Tantousya.Enabled = true;
                    button_Tantousya.BackColor = Color.FromArgb(42, 78, 122);
                }
            }
        }

        private void button_Print_Click(object sender, EventArgs e)
        {
            set_error("", 0);

            if (PrintList.Text == "")
            {
                set_error("", 0);
                set_error("帳票を選択してください。");
            }
            else
            {
                //チェック処理
                Boolean ErrorFlag = false;

                if (TankaKeiyakuID == 0)
                {
                    ErrorFlag = true;
                    // 業務が選択されていませんので、出力できません。
                    set_error(GlobalMethod.GetMessage("E20802", ""));
                }

                if (!ErrorFlag)
                {

                    if (PrintList.SelectedValue.ToString() == "89")
                    {
                        ////業務が選択されていませんので、出力できません。E20802
                        //set_error(GlobalMethod.GetMessage("E20802", "")); //業務(受託番号)はすでに選択済みのため、処理停止

                        //if (TankaKeiyakuID == 0)
                        //{
                        //    ErrorFlag = true;
                        //    // 業務が選択されていませんので、出力できません。
                        //    set_error(GlobalMethod.GetMessage("E20802", ""));
                        //}

                        if (HoukokuSentaku.Text != "")
                        {
                            if (KikanStart.CustomFormat == "" && KikanEnd.CustomFormat == "")
                            {
                                if (KikanStart.Value > KikanEnd.Value)
                                {
                                    ErrorFlag = true;
                                    set_error("期間指定が不正です。");
                                }
                            }
                            else if (KikanStart.CustomFormat != "" && KikanEnd.CustomFormat != "")
                            {
                                ErrorFlag = true;
                                set_error("期間指定を入力してください。");
                            }
                        }

                        //出力処理
                        if (!ErrorFlag)
                        {
                            string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                            using (var conn = new SqlConnection(connStr))
                            {
                                conn.Open();
                                var cmd = conn.CreateCommand();
                                var Dt = new System.Data.DataTable();
                                //SQL生成
                                cmd.CommandText = "SELECT " +
                                  "PrintDataPattern,PrintKikanFlg " +
                                  "FROM " + "Mst_PrintList " +
                                  "WHERE PrintListID = '" + PrintList.SelectedValue + "'";

                                //データ取得
                                var sda = new SqlDataAdapter(cmd);
                                sda.Fill(Dt);
                                //Boolean errorFLG = false;

                                if (Dt.Rows.Count > 0)
                                {
                                    set_error("", 0);
                                    // 10:業務完了内訳表
                                    if (Dt.Rows[0][0].ToString() == "10")
                                    {
                                        // string[]
                                        // 5個分先に用意
                                        string[] report_data = new string[5] { "", "", "", "", "" };

                                        // 0.単価契約ID
                                        report_data[0] = TankaKeiyakuID.ToString();
                                        // 1.日付選択
                                        report_data[1] = "0";
                                        if (HoukokuSentaku.Text != null && HoukokuSentaku.Text != "")
                                        {
                                            report_data[1] = HoukokuSentaku.SelectedValue.ToString();
                                        }
                                        // 2.期間from
                                        report_data[2] = "null";
                                        if (KikanStart.CustomFormat == "")
                                        {
                                            report_data[2] = "'" + KikanStart.Text + "'";
                                        }
                                        // 3.期間to
                                        report_data[3] = "null";
                                        if (KikanEnd.CustomFormat == "")
                                        {
                                            report_data[3] = "'" + KikanEnd.Text + "'";
                                        }
                                        // 4.請求月
                                        report_data[4] = SeikyuuGetsu.Text;

                                        string[] result = GlobalMethod.InsertMadoguchiReportWork(89, UserInfos[0], report_data, "GyoumuKanryou");

                                        // result
                                        // 成否判定 0:正常 1：エラー
                                        // メッセージ（主にエラー用）
                                        // ファイル物理パス（C:\Work\xxxx\0000000111_xxx.xlsx）
                                        // ダウンロード時のファイル名（xxx.xlsx）
                                        if (result != null && result.Length >= 4)
                                        {
                                            if (result[0].Trim() == "1")
                                            {
                                                set_error(result[1]);
                                            }
                                            else
                                            {
                                                Popup_Download form = new Popup_Download();
                                                form.TopLevel = false;
                                                this.Controls.Add(form);

                                                String fileName = Path.GetFileName(result[3]);
                                                form.ExcelName = fileName;
                                                form.TotalFilePath = result[2];
                                                form.Dock = DockStyle.Bottom;
                                                form.Show();
                                                form.BringToFront();
                                            }
                                        }
                                        else
                                        {
                                            // エラーが発生しました
                                            set_error(GlobalMethod.GetMessage("E00091", ""));
                                        }
                                    }
                                }
                                conn.Close();
                            }
                        }
                    }

                    if (PrintList.SelectedValue.ToString() == "229")
                    {

                        DataTable MadoguchiJouhou = new DataTable();
                        using (var conn = new SqlConnection(connStr))
                        {
                            try
                            {
                                conn.Open();
                                var cmd = conn.CreateCommand();

                                // 受託番号から窓口IDを取得する
                                cmd.CommandText = "SELECT TOP 1 MadoguchiID FROM MadoguchiJouhou WHERE MadoguchiJutakuBangou = '" + JutakuBangou + "' ORDER BY MadoguchiID DESC";

                                var sda = new SqlDataAdapter(cmd);
                                MadoguchiJouhou.Clear();
                                sda.Fill(MadoguchiJouhou);

                                if (MadoguchiJouhou.Rows[0][0] != null)
                                {
                                    string MadoguchiID = MadoguchiJouhou.Rows[0][0].ToString();

                                    // 報告書プロンプト
                                    Popup_HoukokuSho form = new Popup_HoukokuSho();
                                    form.MadoguchiID = MadoguchiID;
                                    form.MENU_ID = 208;
                                    form.UserInfos = UserInfos;
                                    form.PrintGamen = "TankaKeiyaku";
                                    if(HoukokuSentaku.Text != null && HoukokuSentaku.Text != "") 
                                    { 
                                        form.HoukokuSentaku = HoukokuSentaku.SelectedValue.ToString();
                                    }
                                    form.KikanStart = NullDate;
                                    if (KikanStart.CustomFormat == "")
                                    {
                                        form.KikanStart = KikanStart.Value;
                                    }
                                    form.KikanEnd = NullDate;
                                    if (KikanEnd.CustomFormat == "")
                                    {
                                        form.KikanEnd = KikanEnd.Value;
                                    }
                                    form.SeikyuuGetsu = SeikyuuGetsu.Text.ToString();
                                    form.ShowDialog();
                                }
                                else
                                {
                                    ErrorFlag = true;
                                    set_error("窓口の登録がありません。");
                                }

                            }
                            catch (Exception)
                            {
                                ErrorFlag = true;
                            }
                            finally
                            {
                                conn.Close();
                            }

                        }

                    }

                }
            }
        }

        private void button_Copy_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            //TankaKeiyakuID = 0;
            Popup_Anken form = new Popup_Anken();
            form.mode = "kakotanka";
            /*
            String discript = "NendoSeireki";
            String value = "NendoID ";
            String table = "Mst_Nendo";
            String where = "Nendo_Sdate <= GETDATE() AND Nendo_EDate >= GETDATE()";
            DataTable dt = GlobalMethod.getData(discript, value, table, where);
            if (dt != null)
            {
                form.nendo = dt.Rows[0][0].ToString();
            }
            else
            {
                form.nendo = DateTime.Today.Year.ToString();
            }
            */
            form.nendo = GlobalMethod.GetTodayNendo();
            //form.hachuushaKaMei = Header_HachuushaKamei.Text.Trim();
            //form.gyoumuMei = Header_GyoumuMei.Text.Trim();
            //form.gyoumuBushoCD = UserInfos[2];
            // 受託部所CDの取得
            DataTable AnkenJouhou = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();

                //単価ランクの取得
                cmd.CommandText = "SELECT AnkenJutakubushoCD FROM AnkenJouhou WHERE AnkenJouhouID = '" + AnkenID + "' ORDER BY AnkenJouhouID";

                var sda = new SqlDataAdapter(cmd);
                AnkenJouhou.Clear();
                sda.Fill(AnkenJouhou);
            }
            form.gyoumuBushoCD = AnkenJouhou.Rows[0][0].ToString();
            //form.gyoumuBushoCD = UserInfos[2];
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                string AnkenJouhouID = form.ReturnValue[0];//案件ID
                int.TryParse(form.ReturnValue[16], out TankaKeiyakuID);//単価契約ID
                TankaKeiyakuID_Copy = TankaKeiyakuID;

                if (AnkenJouhouID != "")
                {
                    get_data(AnkenJouhouID);
                }
            }
        }

        private void button_Ranku_Click(object sender, EventArgs e)
        {
            TankaRankuGrid.Rows.Add();
            TankaRankuGrid.Rows[TankaRankuGrid.Rows.Count - 1][3] = "1";
            Resize_Grid("TankaRankuGrid");
        }

        private void button_KoujiJimusyo_Click(object sender, EventArgs e)
        {
            Popup_KoujiJimusyo1 form = new Popup_KoujiJimusyo1();
            form.data = KoujiJimusyoGrid;
            form.AnkenID = this.AnkenID;
            form.JutakuBangou = Header_JutakuBangou.Text;
            form.ShowDialog();

            if (form.data != null)
            {
                //事務所選択画面から受け取ったデータの表示
                KoujiJimusyoGrid.Rows.Count = 1;
                for (int i = 1; i < form.data.Rows.Count; i++)
                {
                    KoujiJimusyoGrid.Rows.Add();
                    for (int k = 1; k < KoujiJimusyoGrid.Cols.Count; k++)
                    {
                        KoujiJimusyoGrid.Rows[i][k] = form.data.Rows[i][k + 1];
                    }

                    // 工事事務所の行番号に紐づく担当者がいた場合、事務所名を変更する
                    for (int j = 1; j < TantoushaGrid.Rows.Count; j++)
                    {
                        if (TantoushaGrid.Rows[j][8].ToString() == i.ToString())
                        {
                            TantoushaGrid.Rows[j][1] = KoujiJimusyoGrid.Rows[i][2];
                        }
                    }

                }

            }

            Resize_Grid("KoujiJimusyoGrid");
        }

        private void button_Tantousya_Click(object sender, EventArgs e)
        {
            Popup_tantousya form = new Popup_tantousya();
            form.data = TantoushaGrid;
            form.Jimusyodata = KoujiJimusyoGrid;
            form.AnkenID = this.AnkenID;
            form.JutakuBangou = Header_JutakuBangou.Text;
            form.ShowDialog();
            if (form.data != null)
            {
                //担当者選択画面から受け取ったデータの表示
                TantoushaGrid.Rows.Count = 1;
                for (int i = 1; i < form.data.Rows.Count; i++)
                {
                    TantoushaGrid.Rows.Add();
                    for (int k = 1; k < TantoushaGrid.Cols.Count; k++)
                    {
                        TantoushaGrid.Rows[i][k] = form.data.Rows[i][k];
                    }
                }
            }
            Resize_Grid("TantoushaGrid");

            // 工事事務所、部署、担当者名でソートし直す
            TantoushaGrid.Cols[1].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending; // 工事事務所
            TantoushaGrid.Cols[2].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending; // 部署
            TantoushaGrid.Cols[4].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending; // 担当者名

            // 連続した列でないとまとめてソートできないため、並び順を変える
            TantoushaGrid.Cols.Move(4, 3);

            // ソートする
            TantoushaGrid.Sort(C1.Win.C1FlexGrid.SortFlags.UseColSort, 1, 3);

            // 変更した並び順を戻す
            TantoushaGrid.Cols.Move(3, 4);

            // ソート後に△のソートアイコンが出てしまうため、ソート設定をクリアする
            TantoushaGrid.Cols[1].Sort = C1.Win.C1FlexGrid.SortFlags.None; // 工事事務所
            TantoushaGrid.Cols[2].Sort = C1.Win.C1FlexGrid.SortFlags.None; // 部署
            TantoushaGrid.Cols[4].Sort = C1.Win.C1FlexGrid.SortFlags.None; // 担当者名
        }

        private void ShuKeiHouhou_TextChanged(object sender, EventArgs e)
        {
            if (ShuKeiHouhou.Text != "")
            {
                for (int i = 1; i < TankaRankuGrid.Rows.Count; i++)
                {
                    TankaRankuGrid.Rows[i][3] = ShuKeiHouhou.SelectedValue;
                }
            }
        }

        // マウスホイールイベントでコンボ値が変わらないようにする
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        private void dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            ((DateTimePicker)sender).CustomFormat = "";

            DateTime dt = ((DateTimePicker)sender).Value;
            ((DateTimePicker)sender).Text = dt.ToString("yyyy/MM/dd");
        }

        private void dateTimePicker_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Delete)
            {
                ((DateTimePicker)sender).Text = "";
                ((DateTimePicker)sender).CustomFormat = " ";
            }
        }

        private void TankaRankuGrid_BeforeEdit(object sender, RowColEventArgs e)
        {
            switch (e.Col)
            {
                // 価格
                case 2:
                    TankaRankuGrid.ImeMode = ImeMode.Disable;
                    break;
                default:
                    TankaRankuGrid.ImeMode = ImeMode.Off;
                    break;
            }
        }
    }

    public class CPanel : System.Windows.Forms.TableLayoutPanel
    {
        /// <summary>
        /// 自動的にパネル内のオブジェクト位置にスクロールするかどうか
        /// </summary>
        [DefaultValue(false)]
        public bool IsAutoScroll { get; set; }

        /// <summary>
        /// コントロール位置に自動的にスクロールする処理をオーバーライド
        /// </summary>
        /// <param name="c">現在有効なコントロール</param>
        /// <returns>スクロール位置</returns>
        protected override System.Drawing.Point ScrollToControl(Control c)
        {
            if (this.IsAutoScroll)
            {
                return base.ScrollToControl(c);
            }
            else
            {
                return new System.Drawing.Point(-this.HorizontalScroll.Value, -this.VerticalScroll.Value);
            }
        }
    }

}
