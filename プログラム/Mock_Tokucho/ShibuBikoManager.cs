using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using C1.Win.C1FlexGrid;
using System.Configuration;
using System.Data.SqlClient;

namespace TokuchoBugyoK2
{
    class ShibuBikoManager
    {
        public GlobalMethod gm = new GlobalMethod();
        public List<string> chgGrid = new List<String>();

        // ShibuBikou 初期化
        public void ShibuBikoInit(Madoguchi_Input form, Decimal MadoguchiID)
        {
            // データを 'tokuchoBugyoK2DataSet.ShibuBikou' テーブルに読み込みます。
            form.shibuBikouTableAdapter.FillBy(form.tokuchoBugyoKDataSet.ShibuBikou, MadoguchiID);

            // Grid の高さ
            form.BikoGrid.Rows.DefaultSize = 50;
        }

        // ShibuBikou 初期化
        public void ShibuBikoInit(Jibun_Input form, Decimal MadoguchiID)
        {
            // データを 'tokuchoBugyoK2DataSet.ShibuBikou' テーブルに読み込みます。
            form.shibuBikouTableAdapter.FillBy(form.tokuchoBugyoKDataSet.ShibuBikou, MadoguchiID);

            // Grid の高さ
            form.BikoGrid.Rows.DefaultSize = 50;

        }

        public void ShibuBikoInit(Tokumei_Input form, Decimal MadoguchiID)
        {
            // データを 'tokuchoBugyoK2DataSet.ShibuBikou' テーブルに読み込みます。
            form.shibuBikouTableAdapter.FillBy(form.tokuchoBugyoKDataSet.ShibuBikou, MadoguchiID);

            // Grid の高さ
            form.BikoGrid.Rows.DefaultSize = 50;
        }

        // ShibuBikou 更新
        public String UpdateShibuBiko(C1FlexGrid grid, Decimal MadoguchiID)
        {
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            // URL情報更新
            String updsql = "UPDATE dbo.ShibuBikou SET ";
            updsql += " ShibuBikouRyakumei = @Busho";
            updsql += " ,ShibuBikouChousaBusho = @Bushokanri";
            updsql += " ,ShibuBikouKanriNo = @Bushorenban";
            updsql += " ,ShibuBikou = @BushoBikou";
            updsql += " WHERE MadoguchiID = @MadoId ";
            updsql += "  AND ShibuBikouBushoKanriboBushoCD = @Sbkanri ";

            //クエリを実行
            using (SqlConnection con = new SqlConnection(connStr))
            {
                con.Open();
                using (SqlTransaction tran = con.BeginTransaction())
                {
                    try
                    {
                        // ヘッダー込みの為、１スタート
                        for (int row = 1; row < grid.Rows.Count; row++)
                        {
                            Decimal madoId = MadoguchiID; // MadoguchiID
                            String sbkanri = grid.GetData(row, 2).ToString(); // ShibuBikouBushoKanriboBushoCD

                            String busho = grid.GetData(row, 3).ToString(); // 部所
                            String bushokanri = grid.GetData(row, 4).ToString(); // 部所管理
                            String bushoren = grid.GetData(row, 5).ToString(); // 部所連番
                            String bushobiko = grid.GetData(row, 6).ToString(); // 部所備考

                            using (SqlCommand cmd = new SqlCommand(updsql, con, tran))
                            {
                                cmd.Parameters.Add(new SqlParameter("@Busho", busho));
                                cmd.Parameters.Add(new SqlParameter("@Bushokanri", bushokanri));
                                cmd.Parameters.Add(new SqlParameter("@Bushorenban", bushoren));
                                cmd.Parameters.Add(new SqlParameter("@BushoBikou", bushobiko));

                                cmd.Parameters.Add(new SqlParameter("@MadoId", madoId));
                                cmd.Parameters.Add(new SqlParameter("@Sbkanri", sbkanri));

                                cmd.ExecuteNonQuery();
                            }
                        }

                        tran.Commit();
                    }
                    catch (Exception ex)
                    {
                        tran.Rollback();
                        return gm.GetMessage("E20102", "");
                    }
                }
            }
            return gm.GetMessage("I00008", "");
        }

    }
}
