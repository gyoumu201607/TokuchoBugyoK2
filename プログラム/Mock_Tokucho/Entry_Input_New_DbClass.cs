using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using DataTable = System.Data.DataTable;
using System.Collections.Generic;
using System.Linq;

namespace TokuchoBugyoK2
{
    class Entry_Input_New_DbClass
    {
        /// <summary>
        /// DB接続情報
        /// </summary>
        private string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
        private Entry_Input_New form;

        /// <summary>
        /// 計画詳細の「新規登録」で引用あり項目設定処理
        /// </summary>
        public DataTable KeikakuData(string KeikakuID)
        {
            DataTable dtPlan = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                sSql.Append(" SELECT");
                sSql.Append("     KeikakuBangou");
                sSql.Append("    , KeikakuAnkenMei");
                sSql.Append("    , KeikakuZenkaiGyoumuMei");
                sSql.Append("    , KeikakuGyoumuKubun");
                sSql.Append("    , KeikakuKoukiKaishibi");
                sSql.Append("    , KeikakuKoukiShuryoubi");
                sSql.Append("    , KeikakuKaishiNendo");
                sSql.Append("    , KeikakuUriageNendo");
                sSql.Append("    , KeikakuShizaiChousa       AS percent1");
                sSql.Append("    , KeikakuEizen              AS percent2");
                sSql.Append("    , KeikakuKikiruiChousa      AS percent3");
                sSql.Append("    , KeikakuKoujiChousahi      AS percent4");
                sSql.Append("    , KeikakuSanpaiChousa       AS percent5");
                sSql.Append("    , KeikakuHokakeChousa       AS percent6");
                sSql.Append("    , KeikakuShokeihiChousa     AS percent7");
                sSql.Append("    , KeikakuGenkaBunseki       AS percent8");
                sSql.Append("    , KeikakuKijunsakusei       AS percent9");
                sSql.Append("    , KeikakuKoukyouRoumuhi     AS percent10");
                sSql.Append("    , KeikakuRoumuhiKoukyouigai AS percent11");
                sSql.Append("    , KeikakuSonotaChousabu     AS percent12");
                sSql.Append("    , KeikakuHaibunGoukei       AS percentAll");
                sSql.Append("    , CASE WHEN KeikakuMikomigakuGoukei = 0 THEN 0 ELSE KeikakuMikomigaku/KeikakuMikomigakuGoukei*100 END   AS bmPercent1");
                sSql.Append("    , CASE WHEN KeikakuMikomigakuGoukei = 0 THEN 0 ELSE KeikakuMikomigakuJF/KeikakuMikomigakuGoukei*100 END AS bmPercent2");
                sSql.Append("    , CASE WHEN KeikakuMikomigakuGoukei = 0 THEN 0 ELSE KeikakuMikomigakuJ/KeikakuMikomigakuGoukei*100 END  AS bmPercent3");
                sSql.Append("    , CASE WHEN KeikakuMikomigakuGoukei = 0 THEN 0 ELSE KeikakuMikomigakuK/KeikakuMikomigakuGoukei*100 END  AS bmPercent4");
                sSql.Append(" FROM");
                sSql.Append("    KeikakuJouhou");
                sSql.Append(" WHERE KeikakuID = ").Append(KeikakuID);
                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);

                dtPlan.Clear();
                sda.Fill(dtPlan);
            }
            return dtPlan;
        }

        /// <summary>
        /// 案件情報テーブルから取得
        /// </summary>
        /// <param name="AnkenID"></param>
        /// <returns></returns>
        public DataTable AnkenData_H(string AnkenID)
        {
            DataTable dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                sSql.Append("SELECT TOP 1");
                // 基本情報（元引合状況）--------------------------
                sSql.Append("      AnkenJouhou.AnkenJouhouID");
                sSql.Append("    , AnkenSakuseiKubun");
                sSql.Append("    , AnkenUriageNendo");
                sSql.Append("    , AnkenKeikakuBangou");
                sSql.Append("    , KeikakuAnkenMei");
                sSql.Append("    , AnkenAnkenBangou");
                sSql.Append("    , AnkenJutakuBangou");
                sSql.Append("    , AnkenJutakuBangouEda");
                sSql.Append("    , CASE AnkenTourokubi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenTourokubi, 'yyyy/MM/dd') END AS AnkenTourokuDt");
                sSql.Append("    , AnkenJutakubushoCD");
                sSql.Append("    , AnkenTantoushaMei");
                sSql.Append("    , AnkenKeiyakusho");
                //案件情報--------------------------
                sSql.Append("    , AnkenGyoumuMei");
                sSql.Append("    , AnkenGyoumuKubun");
                sSql.Append("    , AnkenNyuusatsuHoushiki");
                sSql.Append("    , CASE AnkenNyuusatsuYoteibi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenNyuusatsuYoteibi, 'yyyy/MM/dd') END AS AnkenNyuusatsuYoteiDt");
                sSql.Append("    , NyuusatsuRakusatsushaID");
                sSql.Append("    , AnkenAnkenMemoKihon");
                //発注者情報--------------------------
                sSql.Append("    , AnkenHachushaCD");
                sSql.Append("    , HachushaKubun1Mei");
                sSql.Append("    , HachushaKubun2Mei");
                sSql.Append("    , TodouhukenMei");
                sSql.Append("    , HachushaMei");
                sSql.Append("    , AnkenHachushaKaMei");
                //発注担当者情報--------------------------
                sSql.Append("    , AnkenHachuushaIraibusho ");
                sSql.Append("    , AnkenHachuushaTantousha ");
                sSql.Append("    , AnkenHachuushaTEL ");
                sSql.Append("    , AnkenHachuushaFAX ");
                sSql.Append("    , AnkenHachuushaMail ");
                sSql.Append("    , AnkenHachuushaIraiYuubin ");
                sSql.Append("    , AnkenHachuushaIraiJuusho ");
                sSql.Append("    , AnkenHachuushaKeiyakuBusho");
                sSql.Append("    , AnkenHachuushaKeiyakuTantou");
                sSql.Append("    , AnkenHachuushaKeiyakuTEL");
                sSql.Append("    , AnkenHachuushaKeiyakuFAX");
                sSql.Append("    , AnkenHachuushaKeiyakuMail");
                sSql.Append("    , AnkenHachuushaKeiyakuYuubin");
                sSql.Append("    , AnkenHachuushaKeiyakuJuusho");
                sSql.Append("    , AnkenHachuuDaihyouYakushoku");
                sSql.Append("    , AnkenHachuuDaihyousha");
                sSql.Append("    , AnkenToukaiSankouMitsumori");
                sSql.Append("    , AnkenToukaiJyutyuIyoku");
                sSql.Append("    , ISNULL(AnkenToukaiSankouMitsumoriGaku, 0)         AS NyuusatsuMitsumoriAmt");
                //部門配分：【事前打診・入札】配分率(%)--------------------------
                sSql.Append("    , ISNULL(GyoumuChosaBuRitsu, 0)             AS BuRitsu1");
                sSql.Append("    , ISNULL(GyoumuJigyoFukyuBuRitsu, 0)        AS BuRitsu2");
                sSql.Append("    , ISNULL(GyoumuJyohouSystemBuRitsu, 0)      AS BuRitsu3");
                sSql.Append("    , ISNULL(GyoumuSougouKenkyuJoRitsu, 0)      AS BuRitsu4");
                //調査部　業務別配分：【事前打診・入札】配分率(%)--------------------------
                sSql.Append("    , ISNULL(GyoumuShizaiChousaRitsu, 0)        AS ChousaRitsu1");
                sSql.Append("    , ISNULL(GyoumuEizenRitsu, 0)               AS ChousaRitsu2");
                sSql.Append("    , ISNULL(GyoumuKikiruiChousaRitsu, 0)       AS ChousaRitsu3");
                sSql.Append("    , ISNULL(GyoumuKoujiChousahiRitsu, 0)       AS ChousaRitsu4");
                sSql.Append("    , ISNULL(GyoumuSanpaiFukusanbutsuRitsu, 0)  AS ChousaRitsu5");
                sSql.Append("    , ISNULL(GyoumuHokakeChousaRitsu, 0)        AS ChousaRitsu6");
                sSql.Append("    , ISNULL(GyoumuShokeihiChousaRitsu, 0)      AS ChousaRitsu7");
                sSql.Append("    , ISNULL(GyoumuGenkaBunsekiRitsu, 0)        AS ChousaRitsu8");
                sSql.Append("    , ISNULL(GyoumuKijunsakuseiRitsu, 0)        AS ChousaRitsu9");
                sSql.Append("    , ISNULL(GyoumuKoukyouRoumuhiRitsu, 0)      AS ChousaRitsu10");
                sSql.Append("    , ISNULL(GyoumuRoumuhiKoukyouigaiRitsu, 0)  AS ChousaRitsu11");
                sSql.Append("    , ISNULL(GyoumuSonotaChousabuRitsu, 0)      AS ChousaRitsu12");

                sSql.Append("    , AnkenKaisuu");
                sSql.Append("    , AnkenSaishinFlg");
                sSql.Append("    , AnkenTantoushaCD");
                sSql.Append("    , Mst_Chousain.GyoumuBushoCD AS AnkenTantoushaBushoCD");
                sSql.Append("    , AnkenKoukiNendo");
                //案件番号変更履歴情報--------------------------
                sSql.Append("    , CASE WHEN AnkenFolderHenkouTantoushaCD = 0 THEN '' ELSE ISNULL(mc.ChousainMei, '') END AS ChousainName");    // 変更者
                sSql.Append("    , ISNULL(AnkenFolderHenkouTantoushaCD, '')  AS AnkenFolderRenameTantouCd");    // 変更者CD
                sSql.Append("    , CASE WHEN AnkenFolderHenkouDatetime IS NULL THEN '' ELSE FORMAT(AnkenFolderHenkouDatetime, 'yyyy/MM/dd HH:mm:ss') END AS AnkenFolderHenkouDt");　//変更日
                sSql.Append("    , ISNULL(AnkenHenkoumaeAnkenBangou, '') AS BefChangeAnkenNo");　//変更前案件番号
                sSql.Append("    , CASE AnkenKeiyakuKoukiKaishibi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenKeiyakuKoukiKaishibi,'yyyy/MM/dd') END AS KeiyakuKoukiKaishibi");
                sSql.Append("    , CASE AnkenKeiyakuKoukiKanryoubi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenKeiyakuKoukiKanryoubi,'yyyy/MM/dd') END AS KeiyakuKoukiKanryoubi");
                sSql.Append("    , ISNULL(mb.JigyoubuHeadCD, '') AS JigyoubuHdCD");
                //進捗階段----------------------------------------------------
                sSql.Append("    , AnkenJizenDashinCheck");
                sSql.Append("    , AnkenJizenDashinDate");
                sSql.Append("    , AnkenNyuusatuCheck");
                sSql.Append("    , AnkenNyuusatuDate");
                sSql.Append("    , AnkenKeiyakuCheck");
                sSql.Append("    , AnkenKeiyakuDate");
                sSql.Append("    , AnkenOueniraiUmu");   // --応援依頼の有無
                sSql.Append("    , AnkenOuenIraiMemo");   // --応援依頼メモ
                sSql.Append("    , AnkenJizenDashinIraibi");   // --事前打診依頼日
                sSql.Append("    , AnkenHachuuYoteiMikomibi");   // --発注予定・見込日
                sSql.Append("    , AnkenMihachuuJoukyou");   // --未発注状況
                sSql.Append("    , AnkenHachuunashiRiyuu");   // --「発注なし」の理由
                sSql.Append("    , AnkenSonotaNaiyou");   // --「その他」の内容
                sSql.Append("    , AnkenAnkenMemoMihachuu");   // --案件メモ(未発注)
                sSql.Append("    , AnkenAnkenMemoJizendashin");   // --案件メモ（事前打診）
                sSql.Append("    , AnkenMihachuuTourokubi");   // --未発注の登録日
                sSql.Append(" FROM");
                sSql.Append("    AnkenJouhou");
                sSql.Append("    LEFT JOIN Mst_Busho mb ");
                sSql.Append("        ON AnkenJutakubushoCD = GyoumuBushoCD ");
                sSql.Append("    LEFT JOIN Mst_Hachusha ");
                sSql.Append("        ON AnkenHachushaCD = HachushaCD ");
                sSql.Append("    LEFT JOIN Mst_HachushaKubun1 ");
                sSql.Append("        ON Mst_HachushaKubun1.HachushaKubun1CD = Mst_Hachusha.HachushaKubun1CD ");
                sSql.Append("    LEFT JOIN Mst_HachushaKubun2 ");
                sSql.Append("        ON Mst_HachushaKubun2.HachushaKubun2CD = Mst_Hachusha.HachushaKubun2CD ");
                sSql.Append("    LEFT JOIN Mst_Todouhuken ");
                sSql.Append("        ON Mst_Todouhuken.TodouhukenCD = Mst_Hachusha.TodouhukenCD ");
                sSql.Append("    LEFT JOIN NyuusatsuJouhou ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN KeikakuJouhou ");
                sSql.Append("        ON AnkenJouhou.AnkenKeikakuBangou = KeikakuJouhou.KeikakuBangou ");
                sSql.Append("    LEFT JOIN GyoumuHaibun ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = GyoumuHaibun.GyoumuAnkenJouhouID ");
                sSql.Append("        AND GyoumuHibunKubun = '10' ");
                sSql.Append("    LEFT JOIN Mst_Chousain ");
                sSql.Append("        ON AnkenJouhou.AnkenTantoushaCD = Mst_Chousain.KojinCD ");
                sSql.Append("    LEFT JOIN Mst_Chousain mc ");
                sSql.Append("        ON AnkenJouhou.AnkenFolderHenkouTantoushaCD = mc.KojinCD ");
                sSql.Append(" WHERE");
                sSql.Append("    AnkenJouhou.AnkenJouhouID = ").Append(AnkenID);
                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
            }
            return dt;
        }

        /// <summary>
        /// 入札（応札）情報テーブル
        /// </summary>
        /// <param name="AnkenID"></param>
        /// <returns></returns>
        public DataTable AnkenData_N(string AnkenID)
        {
            DataTable dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                sSql.Append("SELECT TOP 1");
                sSql.Append("      NyuusatsuRakusatsushaID");
                sSql.Append("    , CASE NyuusatsuRakusatsuKekkaDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(NyuusatsuRakusatsuKekkaDate, 'yyyy/MM/dd') END AS kekkaDate");
                sSql.Append("    , CASE AnkenNyuusatsuYoteibi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenNyuusatsuYoteibi, 'yyyy/MM/dd') END AS yoteibi");
                sSql.Append("    , AnkenHikiaijhokyo");
                sSql.Append("    , AnkenSakuseiKubun");
                sSql.Append("    , NyuusatsuGyoumuBikou");
                sSql.Append("    , AnkenToukaiOusatu");
                sSql.Append("    , AnkenToukaiSankouMitsumori");
                sSql.Append("    , AnkenToukaiJyutyuIyoku");
                sSql.Append("    , ISNULL(NyuusatsuMitsumorigaku, 0) AS NyuusatsuMitsumorigaku");
                sSql.Append("    , NyuusatsuRakusatsuShaJokyou");//10
                sSql.Append("    , NyuusatsuRakusatsuGakuJokyou");
                sSql.Append("    , CASE NyuusatsuRakusatsuShokaiDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(NyuusatsuRakusatsuShokaiDate, 'yyyy/MM/dd') END AS syokaiDate");
                sSql.Append("    , CASE NyuusatsuRakusatsuSaisyuDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(NyuusatsuRakusatsuSaisyuDate, 'yyyy/MM/dd') END AS saisyuDate");
                sSql.Append("    , ISNULL(NyuusatsuYoteiKakaku, 0) AS NyuusatsuYoteiKakaku"); // 14
                sSql.Append("    , ISNULL(NyuusatsushaSuu, 0) AS NyuusatsushaSuu");
                sSql.Append("    , NyuusatsuRakusatsusha");
                sSql.Append("    , ISNULL(NyuusatsuRakusatugaku, 0) AS NyuusatsuRakusatugaku");
                sSql.Append("    , NyuusatsuOusatsusha");
                sSql.Append("    , NyuusatsuOusatsuKingaku");
                sSql.Append("    , NyuusatsuKekkaMemo");
                //業務内容
                sSql.Append("    , ISNULL(GyoumuChosaBuRitsu, 0)        AS BMRitsu1");
                sSql.Append("    , ISNULL(GyoumuJigyoFukyuBuRitsu, 0)   AS BMRitsu2");
                sSql.Append("    , ISNULL(GyoumuJyohouSystemBuRitsu, 0) AS BMRitsu3");
                sSql.Append("    , ISNULL(GyoumuSougouKenkyuJoRitsu, 0) AS BMRitsu4");
                sSql.Append("    , ISNULL(GyoumuShizaiChousaRitsu, 0)       AS GMRitsu1");
                sSql.Append("    , ISNULL(GyoumuEizenRitsu, 0)              AS GMRitsu2");
                sSql.Append("    , ISNULL(GyoumuKikiruiChousaRitsu, 0)      AS GMRitsu3");
                sSql.Append("    , ISNULL(GyoumuKoujiChousahiRitsu, 0)      AS GMRitsu4");
                sSql.Append("    , ISNULL(GyoumuSanpaiFukusanbutsuRitsu, 0) AS GMRitsu5");
                sSql.Append("    , ISNULL(GyoumuHokakeChousaRitsu, 0)       AS GMRitsu6");
                sSql.Append("    , ISNULL(GyoumuShokeihiChousaRitsu, 0)     AS GMRitsu7");
                sSql.Append("    , ISNULL(GyoumuGenkaBunsekiRitsu, 0)       AS GMRitsu8");
                sSql.Append("    , ISNULL(GyoumuKijunsakuseiRitsu, 0)       AS GMRitsu9");
                sSql.Append("    , ISNULL(GyoumuKoukyouRoumuhiRitsu, 0)     AS GMRitsu10");
                sSql.Append("    , ISNULL(GyoumuRoumuhiKoukyouigaiRitsu, 0) AS GMRitsu11");
                sSql.Append("    , ISNULL(GyoumuSonotaChousabuRitsu, 0)     AS GMRitsu12");
                sSql.Append("    , NyuusatsuJouhou.NyuusatsuJouhouID");
                sSql.Append("    , NyuusatsuUpdateDate ");
                //sSql.Append("    , AnkenNyuusatsuHoushiki");
                sSql.Append("    , RIGHT('0' + CONVERT(NVARCHAR, NyuusatsuHoushiki), 2) AS NyuusatsuHoushiki");

                sSql.Append("    , NyuusatsuAnkenMemoNuusatsu");//--案件メモ(入札)
                sSql.Append("    , NyuusatsuSaiitakuSonotaNaiyou");//--その他の内容
                sSql.Append("    , NyuusatsuSaiitakuKinshiNaiyou");//--再委託禁止条項の内容
                sSql.Append("    , NyuusatsuSaiitakuKinshiUmu");//--再委託禁止条項の記載有無
                sSql.Append("    , NyuusatsuJuchuuIyoku");//--受注意欲
                sSql.Append("    , NyuusatsuSankoumitsumoriKingaku");//--参考見積額(税抜)
                sSql.Append("    , NyuusatsuSankoumitsumoriTaiou");//--参考見積対応
                sSql.Append("    , NyuusatsuSaiteiKakakuUmu");//--最低制限価格有無
                sSql.Append("    , NyuusatsuGyoumuHachuukubun");//--業務発注区分
                sSql.Append("    , CASE NyuusatsuJouhouTourokubi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(NyuusatsuJouhouTourokubi, 'yyyy/MM/dd') END AS NyuusatsuJouhouTourokubi");//--入札結果登録日

                sSql.Append(" FROM");
                sSql.Append("    AnkenJouhou ");
                sSql.Append("    LEFT JOIN Mst_SakuseiKubun ");
                sSql.Append("        ON AnkenSakuseiKubun = SakuseiKubunID ");
                sSql.Append("    LEFT JOIN Mst_Busho ");
                sSql.Append("        ON AnkenJutakubushoCD = GyoumuBushoCD ");
                sSql.Append("    LEFT JOIN Mst_KeiyakuKeitai ");
                sSql.Append("        ON AnkenNyuusatsuHoushiki = KeiyakuKeitaiCD ");
                sSql.Append("    LEFT JOIN NyuusatsuJouhou ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN NyuusatsuJouhouOusatsusha ");
                sSql.Append("        ON NyuusatsuJouhou.NyuusatsuJouhouID = NyuusatsuJouhouOusatsusha.NyuusatsuJouhouID ");
                sSql.Append("    LEFT JOIN KeiyakuJouhouEntory ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN GyoumuHaibun ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = GyoumuHaibun.GyoumuAnkenJouhouID ");
                sSql.Append("        AND GyoumuHibunKubun = '20' ");
                sSql.Append(" WHERE");
                sSql.Append("    AnkenJouhou.AnkenJouhouID = ").Append(AnkenID);
                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);

            }
            return dt;
        }

        /// <summary>
        /// 契約情報テーブル
        /// </summary>
        /// <param name="AnkenID"></param>
        /// <returns></returns>
        public DataTable AnkenData_K(string AnkenID)
        {
            DataTable dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                sSql.Append("SELECT TOP 1");
                // 契約タブ
                // 契約情報
                sSql.Append("      AnkenSakuseiKubun ");
                sSql.Append("    , AnkenKianZumi ");
                sSql.Append("    , CASE KeiyakuKeiyakuTeiketsubi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuKeiyakuTeiketsubi,'yyyy/MM/dd') END AS KeiyakuKeiyakuTeiketsubiD");
                sSql.Append("    , CASE KeiyakuSakuseibi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSakuseibi,'yyyy/MM/dd') END AS KeiyakuSakuseibiD");
                sSql.Append("    , AnkenUriageNendo ");
                sSql.Append("    , CASE AnkenKeiyakuKoukiKaishibi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenKeiyakuKoukiKaishibi,'yyyy/MM/dd') END AS KeiyakuKoukiKaishibi");
                sSql.Append("    , CASE AnkenKeiyakuKoukiKanryoubi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenKeiyakuKoukiKanryoubi,'yyyy/MM/dd') END AS KeiyakuKoukiKanryoubi");
                sSql.Append("    , KeiyakuGyoumuKubun ");
                sSql.Append("    , AnkenHachuushaKaMei ");
                sSql.Append("    , KeiyakuShouhizeiritsu ");
                sSql.Append("    , KeiyakuGyoumuMei ");//10
                sSql.Append("    , ISNULL(KeiyakuKeiyakuKingaku,0) AS KeiyakuKeiyakuAmt");
                sSql.Append("    , ISNULL(KeiyakuZeikomiKingaku,0) AS KeiyakuZeikomiAmt");
                sSql.Append("    , ISNULL(KeiyakuuchizeiKingaku,0) AS KeiyakuuchizeiAmt");
                sSql.Append("    , ISNULL(Keiyakukeiyakukingakukei,0) AS KeiyakukeiyakuAmtkukei");
                sSql.Append("    , ISNULL(KeiyakuBetsuKeiyakuKingaku,0) AS KeiyakuBetsuKeiyakuAmt");
                sSql.Append("    , KeiyakuHenkouChuushiRiyuu ");
                sSql.Append("    , NyuusatsuGyoumuBikou ");
                sSql.Append("    , KeiyakuBikou ");
                sSql.Append("    , KeiyakuShosha ");
                sSql.Append("    , KeiyakuTokkiShiyousho ");//20
                sSql.Append("    , KeiyakuMitsumorisho ");
                sSql.Append("    , KeiyakuTanpinChousaMitsumorisho ");
                sSql.Append("    , KeiyakuSonota ");
                sSql.Append("    , KeiyakuSonotaNaiyou ");
                sSql.Append("    , AnkenKeiyakusho ");
                // 配分情報
                sSql.Append("    , ISNULL(KeiyakuUriageHaibunCho,0)    AS Uriage1");
                sSql.Append("    , ISNULL(KeiyakuUriageHaibunJo,0)     AS Uriage2");
                sSql.Append("    , ISNULL(KeiyakuUriageHaibunJosys,0)  AS Uriage3");
                sSql.Append("    , ISNULL(KeiyakuUriageHaibunKei,0)    AS Uriage4");

                sSql.Append("    , ISNULL(GyoumuChosaBuRitsu,0)        AS Haibun1");//30
                sSql.Append("    , ISNULL(GyoumuJigyoFukyuBuRitsu,0)   AS Haibun2");
                sSql.Append("    , ISNULL(GyoumuJyohouSystemBuRitsu,0) AS Haibun3");
                sSql.Append("    , ISNULL(GyoumuSougouKenkyuJoRitsu,0) AS Haibun4");
                // 単契等の見込み補正額
                sSql.Append("    , ISNULL(KeiyakuTankeiMikomiCho,0)    AS Mikomi1");
                sSql.Append("    , ISNULL(KeiyakuTankeiMikomiJo,0)     AS Mikomi2");//35
                sSql.Append("    , ISNULL(KeiyakuTankeiMikomiJosys,0)  AS Mikomi3");
                sSql.Append("    , ISNULL(KeiyakuTankeiMikomiKei,0)    AS Mikomi4");
                // 年度繰越額
                sSql.Append("    , ISNULL(KeiyakuKurikoshiCho,0)       AS Kurikoshi1");
                sSql.Append("    , ISNULL(KeiyakuKurikoshiJo,0)        AS Kurikoshi2");
                sSql.Append("    , ISNULL(KeiyakuKurikoshiJosys,0)     AS Kurikoshi3");//40
                sSql.Append("    , ISNULL(KeiyakuKurikoshiKei,0)       AS Kurikoshi4");
                // 管理者・担当者
                sSql.Append("    , KanriGijutsushaCD ");
                sSql.Append("    , KanriGijutsushaNM ");
                sSql.Append("    , ShousaTantoushaCD ");
                sSql.Append("    , ShousaTantoushaNM ");
                sSql.Append("    , SinsaTantoushaCD ");
                sSql.Append("    , SinsaTantoushaNM ");
                sSql.Append("    , GyoumuKanrishaCD ");
                sSql.Append("    , GyoumuKanrishaMei ");
                sSql.Append("    , GyoumuJouhouMadoKojinCD ");//50
                sSql.Append("    , GyoumuJouhouMadoChousainMei ");
                sSql.Append("    , GyoumuJouhouMadoGyoumuBushoCD ");
                sSql.Append("    , GyoumuJouhouMadoShibuMei ");
                sSql.Append("    , GyoumuJouhouMadoKamei ");
                // 請求書情報
                sSql.Append("    , CASE KeiyakuSeikyuubi1 WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSeikyuubi1,'yyyy/MM/dd') END AS Seikyuubi1");//55
                sSql.Append("    , ISNULL(KeiyakuSeikyuuKingaku1,0) AS SeikyuuAmt1");
                sSql.Append("    , CASE KeiyakuSeikyuubi2 WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSeikyuubi2,'yyyy/MM/dd') END AS Seikyuubi2");
                sSql.Append("    , ISNULL(KeiyakuSeikyuuKingaku2,0) AS SeikyuuAmt2");
                sSql.Append("    , CASE KeiyakuSeikyuubi3 WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSeikyuubi3,'yyyy/MM/dd') END AS Seikyuubi3");
                sSql.Append("    , ISNULL(KeiyakuSeikyuuKingaku3,0) AS SeikyuuAmt3");//60
                sSql.Append("    , CASE KeiyakuSeikyuubi4 WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSeikyuubi4,'yyyy/MM/dd') END AS Seikyuubi4");
                sSql.Append("    , ISNULL(KeiyakuSeikyuuKingaku4,0) AS SeikyuuAmt4");
                sSql.Append("    , CASE KeiyakuSeikyuubi5 WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSeikyuubi5,'yyyy/MM/dd') END AS Seikyuubi5");
                sSql.Append("    , ISNULL(KeiyakuSeikyuuKingaku5,0) AS SeikyuuAmt5");
                sSql.Append("    , CASE KeiyakuZentokinUkewatashibi WHEN '1753/01/01' THEN null ELSE FORMAT(KeiyakuZentokinUkewatashibi,'yyyy/MM/dd') END AS Seikyuubi6");//65
                sSql.Append("    , ISNULL(KeiyakuZentokin,0) AS SeikyuuAmt6");
                // 業務内容
                sSql.Append("    , ISNULL(KeiyakuHaibunChoZeinuki, 0)       AS KeiyakuHaibunZeinuki1");
                sSql.Append("    , ISNULL(KeiyakuHaibunJoZeinuki, 0)        AS KeiyakuHaibunZeinuki2");
                sSql.Append("    , ISNULL(KeiyakuHaibunJosysZeinuki, 0)     AS KeiyakuHaibunZeinuki3");
                sSql.Append("    , ISNULL(KeiyakuHaibunKeiZeinuki, 0)       AS KeiyakuHaibunZeinuki4");//70

                sSql.Append("    , ISNULL(GyoumuShizaiChousaRitsu, 0)       AS GyoumuRitsu1");
                sSql.Append("    , ISNULL(GyoumuEizenRitsu, 0)              AS GyoumuRitsu2");
                sSql.Append("    , ISNULL(GyoumuKikiruiChousaRitsu, 0)      AS GyoumuRitsu3");
                sSql.Append("    , ISNULL(GyoumuKoujiChousahiRitsu, 0)      AS GyoumuRitsu4");
                sSql.Append("    , ISNULL(GyoumuSanpaiFukusanbutsuRitsu, 0) AS GyoumuRitsu5");
                sSql.Append("    , ISNULL(GyoumuHokakeChousaRitsu, 0)       AS GyoumuRitsu6");
                sSql.Append("    , ISNULL(GyoumuShokeihiChousaRitsu, 0)     AS GyoumuRitsu7");
                sSql.Append("    , ISNULL(GyoumuGenkaBunsekiRitsu, 0)       AS GyoumuRitsu8");
                sSql.Append("    , ISNULL(GyoumuKijunsakuseiRitsu, 0)       AS GyoumuRitsu9");
                sSql.Append("    , ISNULL(GyoumuKoukyouRoumuhiRitsu, 0)     AS GyoumuRitsu10");//80
                sSql.Append("    , ISNULL(GyoumuRoumuhiKoukyouigaiRitsu, 0) AS GyoumuRitsu11");
                sSql.Append("    , ISNULL(GyoumuSonotaChousabuRitsu, 0)     AS GyoumuRitsu12");

                sSql.Append("    , ISNULL(GyoumuShizaiChousaGaku, 0)        AS GyoumuGaku1");
                sSql.Append("    , ISNULL(GyoumuEizenGaku, 0)               AS GyoumuGaku2");
                sSql.Append("    , ISNULL(GyoumuKikiruiChousaGaku, 0)       AS GyoumuGaku3");
                sSql.Append("    , ISNULL(GyoumuKoujiChousahiGaku, 0)       AS GyoumuGaku4");
                sSql.Append("    , ISNULL(GyoumuSanpaiFukusanbutsuGaku, 0)  AS GyoumuGaku5");
                sSql.Append("    , ISNULL(GyoumuHokakeChousaGaku, 0)        AS GyoumuGaku6");
                sSql.Append("    , ISNULL(GyoumuShokeihiChousaGaku, 0)      AS GyoumuGaku7");
                sSql.Append("    , ISNULL(GyoumuGenkaBunsekiGaku, 0)        AS GyoumuGaku8");//90
                sSql.Append("    , ISNULL(GyoumuKijunsakuseiGaku, 0)        AS GyoumuGaku9");
                sSql.Append("    , ISNULL(GyoumuKoukyouRoumuhiGaku, 0)      AS GyoumuGaku10");
                sSql.Append("    , ISNULL(GyoumuRoumuhiKoukyouigaiGaku, 0)  AS GyoumuGaku11");
                sSql.Append("    , ISNULL(GyoumuSonotaChousabuGaku, 0)      AS GyoumuGaku12");
                sSql.Append("    , AnkenGyoumuMei");//95
                sSql.Append("    , KeiyakuRIBCYouTankaDataMoushikomisho");
                sSql.Append("    , KeiyakuSashaKeiyu");
                sSql.Append("    , KeiyakuRIBCYouTankaData");
                sSql.Append("    , AnkenKoukiNendo");
                sSql.Append("    , KeiyakuSaiitakuSonotaNaiyou");//--その他の内容
                sSql.Append("    , KeiyakuSaiitakuKinshiNaiyou");//--再委託禁止条項の内容
                sSql.Append("    , KeiyakuSaiitakuKinshiUmu");//--再委託禁止条項の記載有無
                sSql.Append("    , KeiyakuAnkenMemoKeiyaku");//--案件メモ(契約)
                sSql.Append(" FROM");
                sSql.Append("    AnkenJouhou ");
                sSql.Append("    LEFT JOIN NyuusatsuJouhou ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN KeiyakuJouhouEntory ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN GyoumuHaibun ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = GyoumuHaibun.GyoumuAnkenJouhouID ");
                sSql.Append("        AND GyoumuHibunKubun = '30' ");
                sSql.Append("    LEFT JOIN GyoumuJouhou ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = GyoumuJouhou.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN GyoumuJouhouMadoguchi ");
                sSql.Append("        ON GyoumuJouhouMadoguchi.GyoumuJouhouID = GyoumuJouhou.GyoumuJouhouID ");
                sSql.Append(" WHERE");
                sSql.Append("    AnkenJouhou.AnkenJouhouID = ").Append(AnkenID);
                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
            }
            return dt;
        }

        /// <summary>
        /// 業務情報
        /// </summary>
        /// <param name="AnkenID"></param>
        /// <returns></returns>
        public DataTable AnkenData_G(string AnkenID)
        {
            DataTable dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                sSql.Append("SELECT");
                sSql.Append("    GyoumuHyouten");
                sSql.Append("    , KanriGijutsushaNM");
                sSql.Append("    , GyoumuKanriHyouten");
                sSql.Append("    , ShousaTantoushaNM");
                sSql.Append("    , GyoumuShousaHyouten");
                sSql.Append("    , GyoumuTECRISTourokuBangou");
                sSql.Append("    , CASE GyoumuSeikyuubi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(GyoumuSeikyuubi, 'yyyy/MM/dd') END AS GyoumuSeikyuubi");
                sSql.Append("    , AnkenKeiyakusho");
                sSql.Append("    , AnkenKokyakuHyoukaComment");
                sSql.Append("    , AnkenToukaiHyoukaComment ");
                sSql.Append(" FROM");
                sSql.Append("    AnkenJouhou ");
                sSql.Append("    LEFT JOIN GyoumuJouhou ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = GyoumuJouhou.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN NyuusatsuJouhou ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN KeiyakuJouhouEntory ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID ");
                sSql.Append(" WHERE");
                sSql.Append("    AnkenJouhou.AnkenJouhouID = ").Append(AnkenID);

                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
            }
            return dt;
        }

        /// <summary>
        /// 基本情報：過去案件リスト
        /// </summary>
        /// <param name="tokai">マスタ設定値「ENTORY_TOUKAI」のValue2</param>
        /// <param name="AnkenID">案件番号</param>
        /// <param name="isInsert">新規登録か</param>
        /// <returns></returns>
        public DataTable AnkenData_Grid1(string tokai, string AnkenID, bool isInsert = false)
        {
            DataTable dt = new DataTable();
            StringBuilder sSql = new StringBuilder();
            using (var conn = new SqlConnection(connStr)) {
                var cmd = conn.CreateCommand();

                // 共通検索設定
                sSql.Append("SELECT");
                sSql.Append("    AnkenZenkaiAnkenJouhouID");
                sSql.Append("    , AnkenZenkaiAnkenBangou");
                sSql.Append("    , AnkenZenkaiJutakuBangou");
                sSql.Append("    , AnkenZenkaiJutakuEdaban");
                sSql.Append("    , AnkenZenkaiGyoumuMei");
                sSql.Append("    , AnkenZenkaiRakusatsusha");
                sSql.Append("    , AnkenZenkaiRakusatsushaID");

                sSql.Append("    , ISNULL(AnkenZenkaiJutakuKingaku, 0)   AS AnkenZenkaiJutakuKingaku");
                sSql.Append("    , ISNULL(NyuusatsuOusatsuKingaku, 0)    AS NyuusatsuOusatugaku");
                sSql.Append("    , ISNULL(NyuusatsuMitsumorigaku, 0)     AS NyuusatsuMitsumorigaku");

                sSql.Append("    , ISNULL(KeiyakuKeiyakuKingaku, 0)      AS KeiyakuKeiyakuKingaku");
                sSql.Append("    , ISNULL(KeiyakuHaibunZeinukiKei, 0)    AS KeiyakuHaibunZeinukiKei");
                sSql.Append("    , KeiyakuZenkaiRakusatsushaID");
                sSql.Append("    , AnkenZenkaiKyougouKigyouCD");
                sSql.Append("    , AnkenZenkaiRakusatsuID");
                sSql.Append("    , CONCAT(AnkenUriageNendo, '_', CONVERT(NVARCHAR, AnkenNyuusatsuYoteibi, 111), '_', CONVERT(NVARCHAR, AnkenJouhou.AnkenJouhouID)) AS sortKey ");
                sSql.Append(" FROM");
                sSql.Append("    AnkenJouhouZenkaiRakusatsu ");
                sSql.Append("    LEFT JOIN NyuusatsuJouhou ");
                sSql.Append("        ON AnkenJouhouZenkaiRakusatsu.AnkenZenkaiAnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN KeiyakuJouhouEntory ");
                sSql.Append("        ON AnkenJouhouZenkaiRakusatsu.AnkenZenkaiAnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN AnkenJouhou ");
                sSql.Append("        ON AnkenJouhouZenkaiRakusatsu.AnkenZenkaiAnkenJouhouID = AnkenJouhou.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN ( ");
                sSql.Append("        SELECT");
                sSql.Append("            NyuusatsuJouhouID");
                sSql.Append("            , min(NyuusatsuOusatsuKingaku) AS NyuusatsuOusatsuKingaku ");
                sSql.Append("        FROM");
                sSql.Append("            NyuusatsuJouhouOusatsusha ");
                sSql.Append("        WHERE");
                sSql.Append("            NyuusatsuOusatsusha COLLATE Japanese_XJIS_100_CI_AS_SC = N'").Append(tokai).Append("'");
                sSql.Append("        GROUP BY");
                sSql.Append("            NyuusatsuJouhouID");
                sSql.Append("    ) T1 ");
                sSql.Append("        ON T1.NyuusatsuJouhouID = AnkenJouhou.AnkenJouhouID ");
                sSql.Append(" WHERE");
                sSql.Append("    AnkenJouhouZenkaiRakusatsu.AnkenJouhouID = ").Append(AnkenID);
                sSql.Append(" ORDER BY");
                sSql.Append("    AnkenJouhou.AnkenUriageNendo").Append(isInsert ? " DESC" : "");
                sSql.Append("    , AnkenNyuusatsuYoteibi").Append(isInsert ? " DESC" : "");
                sSql.Append("    , AnkenJouhouZenkaiRakusatsu.AnkenJouhouID").Append(isInsert ? " DESC" : "");
                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);

                if (isInsert)
                {
                    // 前回受託番号ID(AnkenZenkaiRakusatsuID)
                    string maxRakusatsuID = "1";
                    if (dt != null) {
                        object max = dt.Compute("MAX(AnkenZenkaiRakusatsuID)", null);
                        maxRakusatsuID = max is DBNull ? "1" : (Convert.ToInt32(max) + 1).ToString();
                    }

                    sSql.Clear();
                    sSql.Append("SELECT");
                    sSql.Append("    AnkenJouhou.AnkenJouhouID               AS AnkenZenkaiAnkenJouhouID");
                    sSql.Append("    , AnkenAnkenBangou                      AS AnkenZenkaiAnkenBangou");
                    sSql.Append("    , AnkenJutakuBangou                     AS AnkenZenkaiJutakuBangou");
                    sSql.Append("    , AnkenJutakuBangouEda                  AS AnkenZenkaiJutakuEdaban");
                    sSql.Append("    , AnkenGyoumuMei                        AS AnkenZenkaiGyoumuMei");
                    sSql.Append("    , NyuusatsuRakusatsusha                 AS AnkenZenkaiRakusatsusha");
                    sSql.Append("    , NyuusatsuRakusatsushaID               AS AnkenZenkaiRakusatsushaID");
                    sSql.Append("    , ISNULL(NyuusatsuRakusatugaku, 0)      AS AnkenZenkaiJutakuKingaku");
                    sSql.Append("    , ISNULL(NyuusatsuOusatsuKingaku, 0)    AS NyuusatsuOusatugaku");
                    sSql.Append("    , NyuusatsuMitsumorigaku                AS NyuusatsuMitsumorigaku");
                    sSql.Append("    , ISNULL(KeiyakuKeiyakuKingaku, 0)      AS KeiyakuKeiyakuKingaku");
                    sSql.Append("    , ISNULL(KeiyakuHaibunZeinukiKei, 0)    AS KeiyakuHaibunZeinukiKei");
                    sSql.Append("    , NyuusatsuKyougouTashaID               AS KeiyakuZenkaiRakusatsushaID");
                    sSql.Append("    , KyougouKigyouCD                       AS AnkenZenkaiKyougouKigyouCD");
                    sSql.Append("    , ").Append(maxRakusatsuID).Append("                                     AS AnkenZenkaiRakusatsuID");
                    sSql.Append("    , CONCAT(AnkenUriageNendo, '_', CONVERT(NVARCHAR, AnkenNyuusatsuYoteibi, 111), '_', CONVERT(NVARCHAR, AnkenJouhou.AnkenJouhouID)) AS sortKey");
                    sSql.Append(" FROM");
                    sSql.Append("    AnkenJouhou ");
                    sSql.Append("    LEFT JOIN NyuusatsuJouhou ");
                    sSql.Append("        ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID ");
                    sSql.Append("    LEFT JOIN KeiyakuJouhouEntory ");
                    sSql.Append("        ON AnkenJouhou.AnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID ");
                    sSql.Append("    LEFT JOIN Mst_KyougouTasha ");
                    sSql.Append("        ON Mst_KyougouTasha.KyougouTashaID = NyuusatsuJouhou.NyuusatsuKyougouTashaID ");
                    sSql.Append("    LEFT JOIN ( ");
                    sSql.Append("        select");
                    sSql.Append("            NyuusatsuJouhouID");
                    sSql.Append("            , min(NyuusatsuOusatsuKingaku) AS NyuusatsuOusatsuKingaku ");
                    sSql.Append("        FROM");
                    sSql.Append("            NyuusatsuJouhouOusatsusha ");
                    sSql.Append("        where");
                    sSql.Append("            NyuusatsuOusatsusha COLLATE Japanese_XJIS_100_CI_AS_SC = N'").Append(tokai).Append("'");
                    sSql.Append("        group by");
                    sSql.Append("            NyuusatsuJouhouID");
                    sSql.Append("    ) T1 ");
                    sSql.Append("        ON T1.NyuusatsuJouhouID = AnkenJouhou.AnkenJouhouID ");
                    sSql.Append(" WHERE");
                    sSql.Append("    AnkenJouhou.AnkenJouhouID = ").Append(AnkenID);
                    cmd.CommandText = sSql.ToString();
                    sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);

                    // 取得した過去案件が5件を超えている場合、追い出しを行う
                    if (dt != null && dt.Rows.Count > 5)
                    {
                        int delCnt = dt.Rows.Count - 5;
                        // 案件前回ID で　昇順でソートする
                        DataRow[] drResult = dt.Select("", "AnkenZenkaiRakusatsuID");
                        DataTable tmpRlt = drResult.CopyToDataTable();
                        for(int i= 0; i < delCnt; i++)
                        {
                            tmpRlt.Rows.RemoveAt(0);
                        }
                        // sortKeyでソート
                        DataRow[] selectedRows = tmpRlt.Select("", "sortKey");

                        dt = new DataTable();
                        // DataRowからDataTableに変換
                        dt = selectedRows.CopyToDataTable();
                    }
                }

            }
            return dt;
        }

        /// <summary>
        /// 入札：入札参加者リスト
        /// </summary>
        /// <param name="AnkenID"></param>
        /// <returns></returns>
        public DataTable AnkenData_Grid2(string AnkenID)
        {
            DataTable dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                sSql.Append("SELECT");
                sSql.Append("    NyuusatsuRakusatsuJyuni");
                sSql.Append("    , NyuusatsuRakusatsuJokyou");
                sSql.Append("    , NyuusatsuOusatsushaID");
                sSql.Append("    , NyuusatsuOusatsusha");
                sSql.Append("    , NyuusatsuOusatsuKingaku");
                sSql.Append("    , NyuusatsuRakusatsuComment");
                sSql.Append("    , NyuusatsuOusatsuKyougouKigyouCD ");
                sSql.Append(" FROM");
                sSql.Append("    AnkenJouhou ");
                sSql.Append("    LEFT JOIN NyuusatsuJouhou ");
                sSql.Append("        ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID ");
                sSql.Append("    LEFT JOIN NyuusatsuJouhouOusatsusha ");
                sSql.Append("        ON NyuusatsuJouhou.NyuusatsuJouhouID = NyuusatsuJouhouOusatsusha.NyuusatsuJouhouID ");
                sSql.Append(" WHERE");
                sSql.Append("    AnkenJouhou.AnkenJouhouID = ").Append(AnkenID);

                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
            }
            return dt;
        }

        /// <summary>
        /// 契約：担当者リスト
        /// </summary>
        /// <param name="AnkenID"></param>
        /// <returns></returns>
        public DataTable AnkenData_Grid3(string AnkenID)
        {
            DataTable dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                
            }
            return dt;
        }

        /// <summary>
        /// 契約：売上計上情報リスト
        /// </summary>
        /// <param name="AnkenID"></param>
        /// <returns></returns>
        public DataTable AnkenData_Grid4(string AnkenID)
        {
            DataTable dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();

                sSql.Append("SELECT");
                sSql.Append("    RibcNo");
                sSql.Append("    , RibcKoukiEnd");
                sSql.Append("    , RibcUriageKeijyoTuki");
                sSql.Append("    , ISNULL(RibcSeikyuKingaku,0)");
                sSql.Append("    , JigyoubuHeadCD");
                sSql.Append("    , RibcKoukiStart");
                sSql.Append("    , RibcNouhinbi ");
                sSql.Append("    , RibcSeikyubi ");
                sSql.Append("    , RibcNyukinyoteibi ");
                sSql.Append("    , RibcKubun ");
                sSql.Append(" FROM");
                sSql.Append("    RibcJouhou ");
                sSql.Append("    LEFT JOIN Mst_Busho ");
                sSql.Append("        ON RibcKankeibusho = GyoumuBushoCD ");
                sSql.Append(" WHERE");
                sSql.Append("    RibcID = ").Append(AnkenID);
                sSql.Append(" ORDER BY RibcKankeibusho, RibcNo");
                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
            }
            return dt;
        }

        /// <summary>
        /// 技術者評価：担当技術者リスト
        /// </summary>
        /// <param name="AnkenID"></param>
        /// <returns></returns>
        public DataTable AnkenData_Grid5(string AnkenID)
        {
            DataTable dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                sSql.Append("SELECT");
                sSql.Append("    HyouronTantoushaCD");
                sSql.Append("    , HyouronTantoushaMei");
                sSql.Append("    , HyouronnTantoushaHyouten");
                sSql.Append(" FROM");
                sSql.Append("    GyoumuJouhouHyouronTantouL1 ");
                sSql.Append(" WHERE");
                sSql.Append("    GyoumuJouhouID = ").Append(AnkenID);
                sSql.Append(" ORDER BY HyouronTantouID");

                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
            }
            return dt;
        }

        /// <summary>
        /// 起案解除処理
        /// </summary>
        /// <returns></returns>
        public bool KianKaijyo(string sAnkenID)
        {
            bool bRtn = false;

            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();

                SqlTransaction transaction = conn.BeginTransaction();
                cmd.Transaction = transaction;

                try
                {
                    cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                "AnkenKianzumi = 0" +
                                " WHERE AnkenJouhouID = " + sAnkenID;
                    cmd.ExecuteNonQuery();

                    transaction.Commit();
                    bRtn = true;
                }
                catch (Exception)
                {
                    transaction.Rollback();
                }
                conn.Close();
            }

            return bRtn;
        }

        /// <summary>
        /// 受託課所支部　種別コード取得（T／J／・・・）
        /// </summary>
        /// <param name=""></param>
        /// <returns></returns>
        public string JigyoubuHeadCD(string sJyutakuKasyoSibuCd)
        {
            string sJigyoubuHeadCD = "";
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                var dt = new System.Data.DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                  "JigyoubuHeadCD " +
                  "FROM " + "Mst_Busho " +
                  "WHERE GyoumuBushoCD = '" + sJyutakuKasyoSibuCd + "' ";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    sJigyoubuHeadCD = dt.Rows[0][0].ToString();
                }
            }
            return sJigyoubuHeadCD;
        }

        #region 案件情報テーブル --------------------------------------------------------------
        /// <summary>
        /// 案件情報テーブルのカラムのリスト
        /// </summary>
        /// <param name="flag"></param>
        /// <returns></returns>
        public string getInsAnkenCols()
        {
            StringBuilder sSql = new StringBuilder();
            sSql.Append("    AnkenJouhouID");
            //sSql.Append("    , AnkenHikiaijhokyo");
            sSql.Append("    , AnkenSakuseiKubun");
            sSql.Append("    , AnkenUriageNendo");
            sSql.Append("    , AnkenKeikakuBangou");
            sSql.Append("    , AnkenAnkenBangou");
            sSql.Append("    , AnkenJutakuBangou");
            sSql.Append("    , AnkenJutakuBangouEda");
            sSql.Append("    , AnkenTourokubi");
            sSql.Append("    , AnkenJutakushibu");
            sSql.Append("    , AnkenJutakubushoCD");
            sSql.Append("    , AnkenKeiyakusho");
            sSql.Append("    , AnkenTantoushaCD");
            sSql.Append("    , AnkenTantoushaMei");
            sSql.Append("    , AnkenGyoumuMei");
            sSql.Append("    , AnkenGyoumuKubun");
            sSql.Append("    , AnkenGyoumuKubunCD");
            sSql.Append("    , AnkenGyoumuKubunMei");
            sSql.Append("    , AnkenNyuusatsuHoushiki");
            sSql.Append("    , AnkenNyuusatsuYoteibi");
            sSql.Append("    , AnkenHachushaCD");
            sSql.Append("    , AnkenHachuushaMei");
            sSql.Append("    , AnkenHachushaKaMei");
            sSql.Append("    , AnkenHachuushaKaMei");


            sSql.Append("    , AnkenHachuushaIraibusho");
            sSql.Append("    , AnkenHachuushaTantousha");
            sSql.Append("    , AnkenHachuushaTEL");
            sSql.Append("    , AnkenHachuushaFAX");
            sSql.Append("    , AnkenHachuushaMail");
            sSql.Append("    , AnkenHachuushaIraiYuubin");
            sSql.Append("    , AnkenHachuushaIraiJuusho");


            sSql.Append("    , AnkenHachuushaKeiyakuBusho");
            sSql.Append("    , AnkenHachuushaKeiyakuTantou");
            sSql.Append("    , AnkenHachuushaKeiyakuTEL");
            sSql.Append("    , AnkenHachuushaKeiyakuFAX");
            sSql.Append("    , AnkenHachuushaKeiyakuMail");
            sSql.Append("    , AnkenHachuushaKeiyakuYuubin");
            sSql.Append("    , AnkenHachuushaKeiyakuJuusho");
            sSql.Append("    , AnkenHachuuDaihyouYakushoku");
            sSql.Append("    , AnkenHachuuDaihyousha");
            sSql.Append("    , AnkenToukaiSankouMitsumori");
            sSql.Append("    , AnkenToukaiJyutyuIyoku");
            sSql.Append("    , AnkenToukaiSankouMitsumoriGaku");
            sSql.Append("    , AnkenCreateProgram");
            sSql.Append("    , AnkenCreateDate");
            sSql.Append("    , AnkenCreateUser");
            sSql.Append("    , AnkenUpdateDate");
            sSql.Append("    , AnkenUpdateUser");
            sSql.Append("    , AnkenDeleteFlag");
            sSql.Append("    , AnkenSaishinFlg");
            sSql.Append("    , AnkenGyoumuKanrishaCD");
            sSql.Append("    , AnkenMadoguchiTantoushaCD");
            sSql.Append("    , GyoumuKanrishaCD");
            sSql.Append("    , AnkenKaisuu");
            sSql.Append("    , AnkenKoukiNendo");
            sSql.Append("    , AnkenKokyakuHyoukaComment");
            sSql.Append("    , AnkenToukaiHyoukaComment");
            sSql.Append("    , AnkenGyoumuKanrisha");
            sSql.Append("    , AnkenKeiyakuKoukiKaishibi");
            sSql.Append("    , AnkenKeiyakuKoukiKanryoubi");
            sSql.Append("    , AnkenKeiyakuDate");      //--契約登録日
            sSql.Append("    , AnkenKeiyakuCheck");     //--契約
            sSql.Append("    , AnkenNyuusatuDate");     //--入札登録日
            sSql.Append("    , AnkenNyuusatuCheck");    //--入札
            sSql.Append("    , AnkenJizenDashinDate");  //--事前打診登録日
            sSql.Append("    , AnkenJizenDashinCheck"); //--事前打診
            sSql.Append("    , AnkenAnkenMemoKihon");   // --案件メモ（基本情報）

            sSql.Append("    , AnkenOueniraiUmu");   // --応援依頼の有無
            sSql.Append("    , AnkenOuenIraiMemo");   // --応援依頼メモ
            sSql.Append("    , AnkenJizenDashinIraibi");   // --事前打診依頼日
            sSql.Append("    , AnkenHachuuYoteiMikomibi");   // --発注予定・見込日
            sSql.Append("    , AnkenMihachuuJoukyou");   // --未発注状況
            sSql.Append("    , AnkenHachuunashiRiyuu");   // --「発注なし」の理由
            sSql.Append("    , AnkenSonotaNaiyou");   // --「その他」の内容
            sSql.Append("    , AnkenToukaiOusatu");   // --当会応札

            return sSql.ToString();
        }
        #endregion

        #region 案件新規作成処理 --------------------------------------------------------------
        /// <summary>
        /// 案件番号取得
        /// </summary>
        /// <param name="sBusyoHdCd"></param>
        /// <param name="sKoukiNendo"></param>
        /// <param name="sBusyoCd"></param>
        /// <returns></returns>
        public string getAnkenNo(string sBusyoHdCd, string sKoukiNendo, string sBusyoCd)
        {
            string ankenNo = "";
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();


                // 業務分類CD + 年度下2桁
                ankenNo = sBusyoHdCd + sKoukiNendo.Substring(2, 2);


                // 課所支部コード
                cmd.CommandText = "SELECT  " +
                        "KashoShibuCD " +

                        //参照テーブル
                        "FROM Mst_Busho " +
                        "WHERE GyoumuBushoCD = '" + sBusyoCd + "' ";
                var sda = new SqlDataAdapter(cmd);
                var dt = new DataTable();
                sda.Fill(dt);

                // 課所支部コードが正しい
                if (dt != null || dt.Rows.Count > 0)
                {
                    ankenNo = ankenNo + dt.Rows[0][0].ToString();
                    cmd.CommandText = "SELECT TOP 1 " +
                            " SUBSTRING(AnkenAnkenBangou,7,3) " +
                            "FROM AnkenJouhou " +
                            "WHERE AnkenAnkenBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'" + ankenNo + "%' and AnkenDeleteFlag != 1 ORDER BY AnkenAnkenBangou DESC";
                    sda = new SqlDataAdapter(cmd);
                    Console.WriteLine(cmd.CommandText);
                    dt = new DataTable();
                    sda.Fill(dt);

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        int AnkenNoRenban;
                        if (int.TryParse(dt.Rows[0][0].ToString(), out AnkenNoRenban))
                        {
                            AnkenNoRenban++;
                        }
                        else
                        {
                            AnkenNoRenban = 1;
                        }
                        ankenNo += string.Format("{0:D3}", AnkenNoRenban);
                    }
                    else
                    {
                        ankenNo += "001";
                    }
                }
                else
                {
                    ankenNo = "";
                }
            }
            return ankenNo;
        }
        #endregion


        #region 案件削除処理 ------------------------------------------------------------------
        /// <summary>
        /// 案件存在チェックなど
        /// </summary>
        /// <param name="sAnkenID"></param>
        /// <param name="iSelf">0:自分検索、1、自分以外検索、2：赤伝検索</param>
        /// <param name="sJyutakuNo"></param>
        /// <param name="sJyutakuEdNo"></param>
        /// <returns></returns>
        public DataTable CheckAnkenBeforeDelete(string sAnkenID, int iSelf = 0, string sJyutakuNo = "", string sJyutakuEdNo = "", SqlConnection cnn = null)
        {
            DataTable dt = new DataTable();
            var conn = cnn;
            if (cnn == null)
            {
                conn = new SqlConnection(connStr);
            }
            using (conn)
            {
                if (cnn == null) conn.Open();
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                sSql.Append("SELECT");
                sSql.Append("    AnkenJouhouID");
                sSql.Append("    , AnkenSaishinFlg");
                sSql.Append("    , AnkenSakuseiKubun");
                sSql.Append("    , AnkenJutakuBangou");
                sSql.Append("    , AnkenJutakuBangouEda ");
                sSql.Append(" FROM");
                sSql.Append("    AnkenJouhou ");
                sSql.Append(" WHERE");
                if (iSelf == 0)
                {
                    // 編集中の案件
                    sSql.Append("    AnkenJouhouID = ").Append(sAnkenID);
                }
                else if (iSelf == 1)
                {
                    // 編集中案件以外
                    sSql.Append("    AnkenJouhouID <> ").Append(sAnkenID);
                    sSql.Append("  AND AnkenJutakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'").Append(sJyutakuNo).Append("'");
                    sSql.Append("  AND AnkenJutakuBangouEda COLLATE Japanese_XJIS_100_CI_AS_SC = N'").Append(sJyutakuEdNo).Append("'");
                    sSql.Append("  AND AnkenSakuseiKubun IN ('01','03','06','07','08','09')");
                    sSql.Append("  AND AnkenJouhou.AnkenDeleteFlag != 1");
                    sSql.Append(" ORDER BY AnkenJutakuBangou DESC, AnkenJouhouID DESC");
                }
                else {
                    // 赤伝検索
                    sSql.Append("    AnkenJutakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'").Append(sJyutakuNo).Append("'");
                    sSql.Append("  AND AnkenJutakuBangouEda COLLATE Japanese_XJIS_100_CI_AS_SC = N'").Append(sJyutakuEdNo).Append("'");
                    sSql.Append("  AND AnkenJouhou.AnkenSakuseiKubun = '02'");
                    sSql.Append("  AND AnkenJouhou.AnkenDeleteFlag != 1");
                    sSql.Append(" ORDER BY AnkenJutakuBangou DESC, AnkenJouhouID DESC");
                }
                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
                if (cnn == null) conn.Close();
            }
            return dt;
        }

        public bool delete(string sAnkenID, string sAnkenSaishinFlg, string sAnkenSakuseiKubun
            , string sJyutakuNo = "", string sJyutakuEdNo = "", string sKeikakuNo = "", string sKeikakuBangou = "")
        {
            bool bRtn = false;
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                SqlTransaction transaction = conn.BeginTransaction();
                var cmd = conn.CreateCommand();
                cmd.Transaction = transaction;

                try
                {
                    cmd.CommandText = "UPDATE KokyakuKeiyakuJouhou SET KokyakuDeleteFlag = 1 " +
                        "WHERE AnkenJouhouID =  " + sAnkenID;
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE GyoumuJouhou SET GyoumuDeleteFlag = 1 " +
                        "WHERE AnkenJouhouID =  " + sAnkenID;
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE KeiyakuJouhouEntory SET KeiyakuDeleteFlag = 1 " +
                        "WHERE AnkenJouhouID =  " + sAnkenID;
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE NyuusatsuJouhou SET NyuusatsuDeleteFlag = 1 " +
                        "WHERE AnkenJouhouID =  " + sAnkenID;
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE AnkenJouhou SET AnkenDeleteFlag = 1 " +
                        "WHERE AnkenJouhouID =  " + sAnkenID;
                    cmd.ExecuteNonQuery();

                    // 最新フラグが0:最新ではない
                    // 契約区分
                    // 03:契約変更（黒伝）
                    // 06:契約変更（黒伝・金額変更）
                    // 07:契約変更（黒伝・工期変更）
                    // 08:契約変更（黒伝・金額工期変更）
                    // 09:契約変更（黒伝・その他）
                    if (sAnkenSaishinFlg == "1" && (sAnkenSakuseiKubun == "03"
                        || sAnkenSakuseiKubun == "06" || sAnkenSakuseiKubun == "07"
                        || sAnkenSakuseiKubun == "08" || sAnkenSakuseiKubun == "09"
                        ))
                    {
                        DataTable dt2 = CheckAnkenBeforeDelete(sAnkenID, 1, sJyutakuNo, sJyutakuEdNo, conn);

                        for (int i = 0; i < dt2.Rows.Count; i++)
                        {
                            if (i == 0)
                            {
                                cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                    " AnkenSaishinFlg = 1 " +
                                    "WHERE AnkenJouhouID =  " + dt2.Rows[i][0].ToString();
                                cmd.ExecuteNonQuery();

                                // 窓口のAnkenJouhouIDも同様に更新する
                                cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                                    "AnkenJouhouID = '" + dt2.Rows[i][0].ToString() + "' " +
                                    ",MadoguchiAnkenJouhouID = '" + dt2.Rows[i][0].ToString() + "' " +
                                    " WHERE MadoguchiJouhou.AnkenJouhouID = " + sAnkenID;
                                cmd.ExecuteNonQuery();

                                // 単価契約のAnkenJouhouIDも同様に更新する
                                cmd.CommandText = "UPDATE TankaKeiyaku SET " +
                                    "AnkenJouhouID = '" + dt2.Rows[i][0].ToString() + "' " +
                                    " WHERE TankaKeiyaku.AnkenJouhouID = " + sAnkenID;
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                    " AnkenSaishinFlg = 0 " +
                                    "WHERE AnkenJouhouID =  " + dt2.Rows[i][0].ToString();
                                cmd.ExecuteNonQuery();
                            }
                        }
                        // 最新フラグを落とす
                        cmd.CommandText = "UPDATE AnkenJouhou SET " +
                        " AnkenSaishinFlg = 0 " +
                        "WHERE AnkenJouhouID = '" + sAnkenID + "'";
                        cmd.ExecuteNonQuery();


                        DataTable dt3 = CheckAnkenBeforeDelete(sAnkenID, 3, sJyutakuNo, sJyutakuEdNo, conn);

                        string akadenAnkenJouhouID = "";
                        // 直近の02:契約変更（赤伝）を削除する
                        if (dt3 != null && dt3.Rows.Count > 0)
                        {
                            akadenAnkenJouhouID = dt3.Rows[0][0].ToString();

                            cmd.CommandText = "UPDATE KokyakuKeiyakuJouhou SET KokyakuDeleteFlag = 1 " +
                                "WHERE AnkenJouhouID =  " + akadenAnkenJouhouID.ToString();
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "UPDATE GyoumuJouhou SET GyoumuDeleteFlag = 1 " +
                                "WHERE AnkenJouhouID =  " + akadenAnkenJouhouID.ToString();
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "UPDATE KeiyakuJouhouEntory SET KeiyakuDeleteFlag = 1 " +
                                "WHERE AnkenJouhouID =  " + akadenAnkenJouhouID.ToString();
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "UPDATE NyuusatsuJouhou SET NyuusatsuDeleteFlag = 1 " +
                                "WHERE AnkenJouhouID =  " + akadenAnkenJouhouID.ToString();
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "UPDATE AnkenJouhou SET AnkenDeleteFlag = 1 " +
                                "WHERE AnkenJouhouID =  " + akadenAnkenJouhouID.ToString();
                            cmd.ExecuteNonQuery();
                        }
                    }
                    // 計画の案件数更新
                    if (sKeikakuNo != "")
                    {
                        cmd.CommandText = "UPDATE KeikakuJouhou SET KeikakuAnkensu = (select count(*) from AnkenJouhou where AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + sKeikakuNo + "' and AnkenDeleteFlag != 1 and AnkenSaishinFlg = 1) WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + sKeikakuNo + "' ";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                    }
                    if (sKeikakuBangou != "")
                    {
                        cmd.CommandText = "UPDATE KeikakuJouhou SET KeikakuAnkensu = (select count(*) from AnkenJouhou where AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + sKeikakuBangou + "' and AnkenDeleteFlag != 1 and AnkenSaishinFlg = 1) WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + sKeikakuBangou + "' ";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                    }

                    transaction.Commit();

                    // 更新履歴の登録
                    //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報を削除しました ID:" + AnkenID, pgmName + methodName, "");

                    bRtn = true;
                    //this.Owner.Show();
                    //this.Close();

                }
                catch (Exception)
                {
                    transaction.Rollback();
                    throw;
                }
            }
            return bRtn;
        }
        #endregion

        #region そのた処理 --------------------------------------------------------------------
        public List<string> getAnkenOuenIraisaki(string ankenID)
        {
            List<string> lst = new List<string>();
            if (string.IsNullOrEmpty(ankenID) == false)
            {
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();
                    var dt = new System.Data.DataTable();

                    StringBuilder sSql = new StringBuilder();
                    sSql.Append("SELECT OueniraisakiCD FROM AnkenOuenIraisaki WHERE AnkenJouhouID = ");
                    sSql.Append(ankenID);

                    //SQL生成
                    cmd.CommandText = sSql.ToString();

                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        lst = dt.AsEnumerable().Select(x => x["OueniraisakiCD"].ToString()).ToList<string>();
                    }
                }
            }
            return lst;
        }
        #endregion
    }
}
