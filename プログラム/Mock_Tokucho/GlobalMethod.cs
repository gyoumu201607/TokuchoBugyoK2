using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using System.Text;
using System.Linq;
using System.ComponentModel;

namespace TokuchoBugyoK2
{
    public class GlobalMethod
    {
        private string pgmName = "GlobalMethod";
        string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

        public Boolean Check_DB()
        {
            try
            {
                //データ取得処理
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                }
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("データベースに接続ができません。" + e.Message);
                return false;
            }
        }

        public string GetPathValid(string Path)
        {
            string ReturnPath = Path;

            string[] SelectedPath = Path.Split('\\');

            for (int i = 0; i < SelectedPath.Length; i++)
            {
                ReturnPath = string.Join("\\", SelectedPath);

                if (Directory.Exists(ReturnPath))
                {
                    break;
                }
                else
                {
                    Array.Resize(ref SelectedPath, SelectedPath.Length - 1);
                }
            }

            return ReturnPath;
        }

        public string[] GetHinagataPath(int PrintID)
        {
            string[] ReturnStrings = new string[2];
            ReturnStrings[0] = GetCommonValue1("HINAGATA_FOLDER");
            var Dt = new DataTable();
            //データ取得処理
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                //SQL生成
                cmd.CommandText = "SELECT PrintFileName,PrintDownloadFileName" +
                    " FROM Mst_PrintList WHERE PrintListID = " + PrintID;
                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(Dt);
            }
            if (Dt.Rows.Count > 0)
            {
                ReturnStrings[0] = ReturnStrings[0] + Dt.Rows[0][0].ToString();
                ReturnStrings[1] = Dt.Rows[0][1].ToString();
            }

            return ReturnStrings;
        }

        public DataTable getData(String DiscrptColum, String valueColum, String table, String where)
        {
            try
            {
                var comboDt = new DataTable();
                //データ取得処理
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    //SQL生成
                    cmd.CommandText = "SELECT " +
                      valueColum + " AS Value ," +
                     DiscrptColum + " AS Discript " +
                     "FROM " + table;
                    //whereがあるときは追加
                    if (where != "")
                    {
                        cmd.CommandText += " WHERE " + where;
                    }
                    Console.WriteLine(cmd.CommandText);
                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(comboDt);
                }
                //データreturn
                return comboDt;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public Boolean Check_Table(string value, string valueColum, string table, string where)
        {
            try
            {
                var comboDt = new DataTable();
                //データ取得処理
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    //SQL生成
                    cmd.CommandText = "SELECT " +
                      valueColum + " AS Value " +
                     "FROM " + table +
                    " WHERE " + valueColum + " = '" + value + "' ";
                    //whereがあるときは追加
                    if (where != "")
                    {
                        cmd.CommandText += " AND " + where;
                    }
                    Console.WriteLine(cmd.CommandText);
                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(comboDt);
                }
                while (comboDt.Rows.Count >= 1)
                {
                    return true;
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // 並び順がKey順
        public SortedList Get_SortedList(DataTable dt)
        {
            SortedList sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            if (dt != null)
            {
                for (int i = 0; dt.Rows.Count > i; i = i + 1)
                {
                    DataRow nr = dt.Rows[i];
                    if (!sl.Contains(nr["Value"].ToString()))
                    {
                        sl.Add(nr["Value"].ToString(), nr["Discript"].ToString());
                    }
                }
            }
            return sl;
        }

        // 並び順が詰めた順
        public ListDictionary Get_ListDictionary(DataTable dt)
        {
            ListDictionary dtMap1 = new ListDictionary();
            //行の数だけの数だけListDictionaryにIDとValueをadd
            if (dt != null)
            {
                for (int i = 0; dt.Rows.Count > i; i = i + 1)
                {
                    DataRow nr = dt.Rows[i];
                    if (!dtMap1.Contains(nr["Value"].ToString()))
                    {
                        dtMap1.Add(nr["Value"].ToString(), nr["Discript"].ToString());
                    }
                }
            }
            return dtMap1;
        }

        public void Get_WorkFolder()
        {
            string path = System.IO.Path.Combine(new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).Directory.FullName, @"Work");
            if (!File.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        // サニタイジング処理
        // flg:0 なら ' のみ対応、0以外なら複数対応（金額、%表示の項目等に使える）
        // lengh が指定されている場合、指定した文字数までにsubstringする
        public string ChangeSqlText(string str, int flg, int length = 0)
        {
            if (length != 0 && str.Length > length)
            {
                str = str.Substring(0, length);
            }
            if (flg == 0)
            {
                str = str.Replace("'", "''");
            }
            else
            {
                str = str.Replace("'", "''");
                str = str.Replace("%", @"\%");
                str = str.Replace("_", @"\_");
                str = str.Replace("[", @"\[");
                str = str.Replace("]", @"\]");
                str = str.Replace("^", @"\^");
                str = str.Replace(@"\", @"\\");
            }

            return str;
        }

        public DataTable Check_Login(string ID, string Pass)
        {
            using (var conn = new SqlConnection(connStr))
            {
                var dt = new DataTable();
                var cmd = conn.CreateCommand();
                //SQL生成
                cmd.CommandText = "SELECT " +
                 "M_USER.USER_KojinCD,USER_MEI,ROLE_ID " +
                 "FROM M_USER " +
                 "LEFT JOIN M_USERROLE ON M_USER.USER_ID = M_USERROLE.USER_ID " +
                 "WHERE M_USER.USER_ID = '" + ChangeSqlText(ID, 0, 0) + "' AND USER_PASSWORD = '" + ChangeSqlText(Pass, 0, 0) + "'";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
                if (dt.Rows.Count == 0)
                {
                    return null;
                }
                return dt;
            }
        }

        public DataTable Check_Login_Chousain(string KojinCD, string mail)
        {
            using (var conn = new SqlConnection(connStr))
            {
                string FromNendo = DateTime.Now.Year.ToString();
                string ToNendo = (int.Parse(FromNendo) + 1).ToString();
                var dt = new DataTable();
                var cmd = conn.CreateCommand();
                //SQL生成
                cmd.CommandText = "SELECT " +
                 "KojinCD,ChousainMei,Mst_Busho.GyoumuBushoCD,Mst_Busho.ShibuMei + IsNull(Mst_Busho.KaMei,''),TokuchoRole " +
                 "FROM Mst_Chousain " +
                 "LEFT JOIN Mst_Busho ON Mst_Chousain.GyoumuBushoCD = Mst_Busho.GyoumuBushoCD " +
                 "WHERE RetireFLG = 0 AND TokuchoFLG >= 1 AND ISNULL(ChousainDeleteFlag,0) = 0 " +
                 //"AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                 //"AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                 "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today.ToString() + "') " +
                 "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today.ToString() + "' ) ";
                if (KojinCD == "")
                {
                    cmd.CommandText += "AND ChousainID = '" + mail + "' ";
                }
                else
                {
                    cmd.CommandText += "AND KojinCD = '" + KojinCD + "' ";
                }

                //データ取得
                //Console.WriteLine(cmd.CommandText);
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
                if (dt.Rows.Count == 0)
                {
                    return null;
                }
                return dt;
            }
        }

        public DataTable getError(string errorId, int Count, string Where, string chousaHinmokuErrorCnt = "", string FileReadErrorTokuchoBangou = "")
        {
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                var comboDt = new System.Data.DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                "FileReadErrorLineNo ,FileReadErrorMessage,LTRIM(RTRIM(FileReadErrorFilename)),FileReadErrorDateTime " +
                "FROM " + "T_FileReadError " +
                "WHERE ((FileReadErrorTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + errorId + "' " +
                " AND FileReadErrorReadCount = " + Count;
                cmd.CommandText += ") ";

                if (chousaHinmokuErrorCnt != "" && FileReadErrorTokuchoBangou != "")
                {
                    // 窓口ミハル一括登録で、調査品目のエラーも出す
                    cmd.CommandText += " OR (FileReadErrorTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + FileReadErrorTokuchoBangou + "' AND FileReadErrorReadCount = " + chousaHinmokuErrorCnt + " ) ";
                }

                cmd.CommandText += ") ";
                if (Where != ")")
                {
                    cmd.CommandText += " AND FileReadErrorMessage COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + ChangeSqlText(Where, 1, 0) + "%'  ESCAPE '\\' ";
                }

                //なぜかコニカミノルタ環境だと取得するDBの順序が、文字ソートみたいになっているため、order by を明示する
                cmd.CommandText += " ORDER BY FileReadErrorID ASC , FileReadErrorLineNo ASC";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(comboDt);

                //MessageBox.Show(comboDt.Rows.Count + "テスト", "test", MessageBoxButtons.OK);
                //データセット
                return comboDt;
            }
        }
        //課題No1300（994）VIPS Overload
        public DataTable getError(string errorId, List<int> Counts, string Where, string chousaHinmokuErrorCnt = "", string FileReadErrorTokuchoBangou = "")
        {
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                var comboDt = new System.Data.DataTable();
                //FileReadErrorReadCountのin文字列を作成する
                string inBuff = "";
                foreach(int num in Counts)
                {
                    if (inBuff != "")
                    {
                        inBuff += ",";
                    }
                    inBuff += num.ToString();
                }
                //エラーがすべてファイルが見つからないなどの場合DBに登録されないのでゼロで検索する（もともとの仕様に合わせる）
                if (inBuff == "")
                {
                    inBuff = " = 0";
                }
                else
                {
                    inBuff = " in (" + inBuff + ")";
                }
                
                //SQL生成
                cmd.CommandText = "SELECT " +
                "FileReadErrorLineNo ,FileReadErrorMessage,LTRIM(RTRIM(FileReadErrorFilename)),FileReadErrorDateTime " +
                "FROM " + "T_FileReadError " +
                "WHERE ((FileReadErrorTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + errorId + "' " +
                " AND FileReadErrorReadCount " + inBuff;
                cmd.CommandText += ") ";

                if (chousaHinmokuErrorCnt != "" && FileReadErrorTokuchoBangou != "")
                {
                    // 窓口ミハル一括登録で、調査品目のエラーも出す
                    cmd.CommandText += " OR (FileReadErrorTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + FileReadErrorTokuchoBangou + "' AND FileReadErrorReadCount = " + chousaHinmokuErrorCnt + " ) ";
                }

                cmd.CommandText += ") ";
                if (Where != ")")
                {
                    cmd.CommandText += " AND FileReadErrorMessage COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + ChangeSqlText(Where, 1, 0) + "%'  ESCAPE '\\' ";
                }

                //なぜかコニカミノルタ環境だと取得するDBの順序が、文字ソートみたいになっているため、order by を明示する
                cmd.CommandText += " ORDER BY FileReadErrorID ASC , FileReadErrorLineNo ASC";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(comboDt);

                //MessageBox.Show(comboDt.Rows.Count + "テスト", "test", MessageBoxButtons.OK);
                //データセット
                return comboDt;
            }
        }

        public void outputLogger(String programName, String logTitle, String avalable, String logUser, int mode = 1)
        {
            var comboDt = new DataTable();
            //データ取得処理
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                //デバッグモードの確認　コモンマスタ
                //0=運用モード、1=デバッグモード
                var cmd = conn.CreateCommand();

                //デバッグの場合
                if (CheckDebugMode(mode))
                {
                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;
                    try
                    {
                        //デバッグテーブルへ書き込み
                        cmd.CommandText = "INSERT INTO T_DEBUGLOG(" +
                        "EditLogFunctionName , EditlogDitail , EditLogAvalable," +
                        "EditLogCreateDateTime , EditLogUser) " +
                        " VALUES ( '" + programName + "' ,N'" + logTitle + "' ,N'" + avalable +
                        "' ,SYSDATETIME(),N'" + logUser + "')";

                        cmd.ExecuteNonQuery();
                        transaction.Commit();
                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                    }
                }
                conn.Close();
            }
        }

        public DialogResult outputMessage(String messageId, String specified, int flg = 0)
        {
            if (flg == 0)
            {
                return MessageBox.Show(GetMessage(messageId, specified), "確認", MessageBoxButtons.OK);
            }
            else
            {
                return MessageBox.Show(GetMessage(messageId, specified), "確認", MessageBoxButtons.OKCancel);
            }
        }

        public string GetMessage(String messageId, String specified)
        {
            //M_Messageからメッセージを取得する
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                var comboDt = new System.Data.DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                  "Message " +
                  "FROM " + "M_Message " +
                  "WHERE MessageID = '" + messageId + "'";



                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(comboDt);



                String message = "";
                if (comboDt.Rows.Count > 0)
                {
                    message = comboDt.Rows[0][0].ToString();
                }
                //　メッセージの後ろに追加文字列がある場合
                if (specified != "")
                {
                    message += "(" + specified + ")";
                }
                conn.Close();
                if (message == "")
                {
                    message = "メッセージが取得できませんでした。（" + messageId + "）";
                }
                return message;
            }
        }

        public void CreateFolder(int ankenID)
        {
            //string BasePath = GetCommonValue1("FOLDER_BASE").Replace("/","\\");
            string BasePath = "";
            string KashoCD = "";
            string JigyoubuHeadCD = "";
            string BaseFile = "";

            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                var dt1 = new System.Data.DataTable();
                cmd.CommandText = "SELECT " +
                  //"AnkenUriageNendo " +
                  "AnkenKoukiNendo " +
                  ",KashoShibuCD " +
                  ",AnkenAnkenBangou " +
                  ",ISNULL(HachushaMei,'') " +
                  ",ISNULL(AnkenGyoumuMei,'') " +
                  ",AnkenKeiyakusho " +
                  ",JigyoubuHeadCD " +
                  "FROM AnkenJouhou " +
                  "LEFT JOIN Mst_Busho ON GyoumuBushoCD = AnkenJutakubushoCD " +
                  "LEFT JOIN Mst_Hachusha ON HachushaCD = AnkenHachushaCD " +

                //"WHERE GyoumuBushoCD = '" + ankenID + "'";
                "WHERE AnkenJouhouID = '" + ankenID + "'";
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt1);
                conn.Close();

                if (dt1 != null && dt1.Rows.Count > 0)
                {
                    string Nendo = "";
                    Nendo = dt1.Rows[0][0].ToString();
                    KashoCD = dt1.Rows[0][1].ToString();
                    BasePath = dt1.Rows[0][5].ToString();
                    JigyoubuHeadCD = dt1.Rows[0][6].ToString();

                    // 調査部 JigyoubuHeadCD:T以外（他部所）はフォルダは作成しない
                    if (!"T".Equals(JigyoubuHeadCD))
                    {
                        return;
                    }

                    // 案件（受託）フォルダーを取得し、/を\に置換する
                    BasePath = BasePath.Replace("/", "\\");

                    string AnkenBangou = dt1.Rows[0][2].ToString();
                    string Hachusha = dt1.Rows[0][3].ToString();
                    // 空白はトリム
                    Hachusha = System.Text.RegularExpressions.Regex.Replace(Hachusha, @"\s", "");
                    if (Hachusha.Length > 10)
                    {
                        Hachusha = Hachusha.Substring(0, 10);
                    }
                    string AnkenMei = dt1.Rows[0][4].ToString();
                    AnkenMei = System.Text.RegularExpressions.Regex.Replace(AnkenMei, @"\s", "");
                    if (AnkenMei.Length > 20)
                    {
                        AnkenMei = AnkenMei.Substring(0, 20);
                    }
                    BaseFile = AnkenBangou + "_" + Hachusha + "_" + AnkenMei;
                }
            }
            try
            {
                // フォルダ存在フラグ true：存在する false：存在しない
                Boolean ExistFlg = false;
                if (!"".Equals(BasePath))
                {
                    string[] fileList = Directory.GetFileSystemEntries(BasePath, KashoCD + "*");
                    //foreach (var filePath in fileList)
                    //{
                    //    BasePath = filePath;
                    //    ExistFlg = true;
                    //    break;
                    //}

                    // 案件（受託）フォルダで指定されたパスが存在するか
                    if (Directory.Exists(BasePath))
                    {
                        ExistFlg = true;
                    }
                }

                if (!ExistFlg)
                {
                    // 支部のフォルダが見つかりませんでした。
                    outputMessage("E70047", "");
                    return;
                }
            }
            catch (Exception)
            {
                return;
            }

            DataTable CreateList = getData("CommonMasterID", "CommonValue1", "M_CommonMaster", "CommonMasterKye = 'ANKEN_BANGOU_FOLDER' ORDER BY CommonMasterID ");
            BasePath += "\\" + BaseFile;

            // 案件受託フォルダ + AnkenBangou + "_" + Hachusha + "_" + AnkenMei が存在するか
            if (!File.Exists(BasePath))
            {
                try
                {
                    DirectoryInfo di = new DirectoryInfo(BasePath);
                    di.Create();
                    //File.Create(BasePath);
                }
                catch (Exception)
                {
                    // フォルダを作成する権限がありません。
                    outputMessage("E70046", "");
                    return;
                }
            }

            // CommonMasterにANKEN_BANGOU_FOLDERが存在するかどうか
            if (CreateList.Rows.Count > 0)
            {
                for (int i = 0; i < CreateList.Rows.Count; i++)
                {

                    DirectoryInfo di = new DirectoryInfo(BasePath + "\\" + CreateList.Rows[i][0].ToString());

                    //if (!File.Exists(BasePath + "\\" + CreateList.Rows[i][0].ToString()))
                    if (!Directory.Exists(BasePath + "\\" + CreateList.Rows[i][0].ToString()))
                    {
                        try
                        {
                            di.Create();
                            //File.Create(BasePath + "\\" + CreateList.Rows[i][0].ToString());
                        }
                        catch (Exception)
                        {
                            // フォルダを作成する権限がありません。
                            outputMessage("E70046", "");
                            return;
                        }
                    }
                }
                // ここまできたらフォルダ作成が成功なので、
                // 案件（受託）フォルダに作成したAnkenBangou + "_" + Hachusha + "_" + AnkenMeiを付与し更新
                using (var conn = new SqlConnection(connStr))
                {
                    // ファルダ生成成功時
                    conn.Open();
                    var cmd = conn.CreateCommand();
                    var dt1 = new System.Data.DataTable();
                    cmd.CommandText = "UPDATE AnkenJouhou SET " +
                    "AnkenKeiyakusho = AnkenKeiyakusho + " + "N'" + @"\" + ChangeSqlText(BaseFile, 0, 0) + "' " +
                    "WHERE AnkenJouhouID = '" + ankenID + "'";
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt1);
                    conn.Close();
                }
            }

        }

        // 窓口新規登録時に
        // 04集計表
        // 05報告書
        // 06調査資料・図面
        // 配下に、特調枝番フォルダを作成する
        public string CreateTokuchoBangouEdaFolder(string MadoguchiID, string tokuchoBangouEda)
        {
            string message = "";
            string JigyoubuHeadCD = "";
            string MadoguchiShukeiHyoFolder = "";
            string MadoguchiHoukokuShoFolder = "";
            string MadoguchiShiryouHolder = "";
            string AnkenUriageNendo = "";

            if (MadoguchiID == "")
            {
                return message;
            }

            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                var dt1 = new System.Data.DataTable();
                cmd.CommandText = "SELECT TOP 1 " +
                  "JigyoubuHeadCD " +
                  ",MadoguchiShukeiHyoFolder " +  // 集計表
                  ",MadoguchiHoukokuShoFolder " + // 報告書
                  ",MadoguchiShiryouHolder " +    // 調査資料
                                                  //",AnkenUriageNendo " +          // 売上年度
                  ",AnkenKoukiNendo " +          // 工期開始年度
                  "FROM MadoguchiJouhou " +
                  "LEFT JOIN Mst_Busho ON GyoumuBushoCD = MadoguchiJutakuBushoCD " +
                  //"LEFT JOIN AnkenJouhou ON AnkenJutakuBangou = replace(MadoguchiJutakuBangou,'-' + MadoguchiJutakuBangouEdaban,'') AND AnkenDeleteFlag != 1 AND AnkenSaishinFlg = 1 " +
                  "LEFT JOIN AnkenJouhou ON AnkenJutakuBangou = MadoguchiJutakuBangou + '-' + MadoguchiJutakuBangouEdaban AND AnkenDeleteFlag != 1 AND AnkenSaishinFlg = 1 " +
                "WHERE MadoguchiDeleteFlag != 1 AND MadoguchiID = " + MadoguchiID + "";

                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt1);
                conn.Close();

                if (dt1 != null && dt1.Rows.Count > 0)
                {
                    JigyoubuHeadCD = dt1.Rows[0][0].ToString();
                    MadoguchiShukeiHyoFolder = dt1.Rows[0][1].ToString();
                    MadoguchiHoukokuShoFolder = dt1.Rows[0][2].ToString();
                    MadoguchiShiryouHolder = dt1.Rows[0][3].ToString();
                    AnkenUriageNendo = dt1.Rows[0][4].ToString();

                    // 調査部 JigyoubuHeadCD:T以外（他部所）はフォルダは作成しない
                    if (!"T".Equals(JigyoubuHeadCD))
                    {
                        return message;
                    }

                    if (AnkenUriageNendo != "")
                    {
                        int nendo = int.Parse(AnkenUriageNendo);
                        // 2021年より前はフォルダを作成しない
                        if (nendo < 2021)
                        {
                            return message;
                        }
                    }

                    try
                    {
                        string ShukeiHyoFolder = GetCommonValue1("ANKEN_BANGOU_FOLDER", "13"); // 集計表
                        string HoukokuShoFolder = GetCommonValue1("ANKEN_BANGOU_FOLDER", "14"); // 報告書
                        string ShiryouHolder = GetCommonValue1("ANKEN_BANGOU_FOLDER", "15"); // 調査資料・図面

                        // フォルダ存在フラグ true：存在する false：存在しない
                        Boolean ShukeiExistFlg = false;
                        Boolean HoukokuExistFlg = false;
                        Boolean ShiryouExistFlg = false;
                        // 集計表フォルダ確認
                        if (!"".Equals(MadoguchiShukeiHyoFolder) && MadoguchiShukeiHyoFolder.IndexOf(ShukeiHyoFolder) >= 0)
                        {
                            // フォルダで指定されたパスが存在するか
                            if (Directory.Exists(MadoguchiShukeiHyoFolder))
                            {
                                ShukeiExistFlg = true;
                            }
                        }
                        if (!ShukeiExistFlg)
                        {
                            // 集計表フォルダが見つかりませんでした。
                            //outputMessage("E70050", "");
                            message = GetMessage("E70050", "") + Environment.NewLine;
                            return message;
                        }
                        // 報告書フォルダ確認
                        if (!"".Equals(MadoguchiHoukokuShoFolder) && MadoguchiHoukokuShoFolder.IndexOf(HoukokuShoFolder) >= 0)
                        {
                            // フォルダで指定されたパスが存在するか
                            if (Directory.Exists(MadoguchiShukeiHyoFolder))
                            {
                                HoukokuExistFlg = true;
                            }
                        }
                        if (!HoukokuExistFlg)
                        {
                            // 報告書フォルダが見つかりませんでした。
                            //outputMessage("E70051", "");
                            message = GetMessage("E70051", "") + Environment.NewLine;
                            return message;
                        }
                        // 調査資料・図面フォルダ確認
                        if (!"".Equals(MadoguchiShiryouHolder) && MadoguchiShiryouHolder.IndexOf(ShiryouHolder) >= 0)
                        {
                            // フォルダで指定されたパスが存在するか
                            if (Directory.Exists(MadoguchiShiryouHolder))
                            {
                                ShiryouExistFlg = true;
                            }
                        }
                        if (!ShiryouExistFlg)
                        {
                            // 調査資料・図面フォルダが見つかりませんでした。
                            //outputMessage("E70052", "");
                            message = GetMessage("E70052", "") + Environment.NewLine;
                            return message;
                        }

                        // 集計表フォルダ作成
                        try
                        {
                            MadoguchiShukeiHyoFolder = MadoguchiShukeiHyoFolder + @"\" + tokuchoBangouEda;
                            if (!Directory.Exists(MadoguchiShukeiHyoFolder))
                            {
                                DirectoryInfo di = new DirectoryInfo(MadoguchiShukeiHyoFolder);
                                di.Create();
                            }
                        }
                        catch (Exception)
                        {
                            // フォルダを作成する権限がありません。
                            //outputMessage("E70046", "(集計表)");
                            message = GetMessage("E70046", "(集計表)") + Environment.NewLine;
                            return message;
                        }
                        // 報告書フォルダ作成
                        try
                        {
                            MadoguchiHoukokuShoFolder = MadoguchiHoukokuShoFolder + @"\" + tokuchoBangouEda;
                            if (!Directory.Exists(MadoguchiHoukokuShoFolder))
                            {
                                DirectoryInfo di = new DirectoryInfo(MadoguchiHoukokuShoFolder);
                                di.Create();
                            }
                        }
                        catch (Exception)
                        {
                            // フォルダを作成する権限がありません。
                            //outputMessage("E70046", "(報告書)");
                            message = GetMessage("E70046", "(報告書)") + Environment.NewLine;
                            return message;
                        }
                        // 調査資料・図面フォルダ作成
                        try
                        {
                            MadoguchiShiryouHolder = MadoguchiShiryouHolder + @"\" + tokuchoBangouEda;
                            if (!Directory.Exists(MadoguchiShiryouHolder))
                            {
                                DirectoryInfo di = new DirectoryInfo(MadoguchiShiryouHolder);
                                di.Create();
                            }
                        }
                        catch (Exception)
                        {
                            // フォルダを作成する権限がありません。
                            //outputMessage("E70046", "(調査資料・図面)");
                            message = GetMessage("E70046", "(調査資料・図面)") + Environment.NewLine;
                            return message;
                        }

                        // ファルダ生成成功時
                        conn.Open();
                        var cmd2 = conn.CreateCommand();
                        var dt2 = new System.Data.DataTable();
                        cmd2.CommandText = "UPDATE MadoguchiJouhou SET " +
                        "MadoguchiShukeiHyoFolder = N'" + ChangeSqlText(MadoguchiShukeiHyoFolder, 0, 0) + "'" + // 集計表
                        ",MadoguchiHoukokuShoFolder = N'" + ChangeSqlText(MadoguchiHoukokuShoFolder, 0, 0) + "'" + // 報告書
                        ",MadoguchiShiryouHolder = N'" + ChangeSqlText(MadoguchiShiryouHolder, 0, 0) + "' " + // 調査資料
                        "WHERE MadoguchiID = " + MadoguchiID + " ";
                        var sda2 = new SqlDataAdapter(cmd2);
                        sda2.Fill(dt2);
                        conn.Close();
                    }
                    catch (Exception)
                    {
                        // フォルダ作成時にエラーが発生しました
                        message = GetMessage("E70053", "");
                        return "";
                    }
                    // 成功時にはなにも付与しない
                    return "";
                }
                return "";
            }
        }

        public bool CheckDebugMode(int mode = 1)
        {
            try
            {
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    string sqlstr = "select CommonValue1 from M_CommonMaster where CommonMasterKye = 'DEBUG_MODE'";
                    SqlCommand com = new SqlCommand(sqlstr, conn);
                    SqlDataReader sdr = com.ExecuteReader();
                    string value = "";
                    while (sdr.Read() == true)
                    {
                        value = (string)sdr["CommonValue1"];
                    }
                    return value == mode.ToString();
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                return false;
            }
        }

        public int GetIntroductionPhase()
        {
            int num = 9;
            SqlConnection sqlconn = new SqlConnection(connStr);
            sqlconn.Open();
            try
            {
                string sqlstr = "select CommonValue1 from M_CommonMaster where CommonMasterKye = 'INTRODUCTION_PHASE_FLAG'";
                SqlCommand com = new SqlCommand(sqlstr, sqlconn);
                SqlDataReader sdr = com.ExecuteReader();

                while (sdr.Read() == true)
                {
                    num = int.Parse((string)sdr["CommonValue1"]);
                }

            }
            finally
            {
                sqlconn.Close();
            }


            return num;
        }

        public string GetCommonValue1(string Key, string ID = "1")
        {
            string output = null;
            SqlConnection sqlconn = new SqlConnection(connStr);
            sqlconn.Open();
            try
            {
                string sqlstr = "select CommonValue1 from M_CommonMaster where CommonMasterKye = '" + Key + "' AND CommonMasterID = '" + ID + "' ";
                SqlCommand com = new SqlCommand(sqlstr, sqlconn);
                SqlDataReader sdr = com.ExecuteReader();

                while (sdr.Read() == true)
                {
                    output = (string)sdr["CommonValue1"];
                }
            }
            finally
            {
                sqlconn.Close();
            }
            return output;
        }

        public string GetCommonValue2(string Key)
        {
            string output = null;
            SqlConnection sqlconn = new SqlConnection(connStr);
            sqlconn.Open();
            try
            {
                string sqlstr = "select CommonValue2 from M_CommonMaster where CommonMasterKye = '" + Key + "'";
                SqlCommand com = new SqlCommand(sqlstr, sqlconn);
                SqlDataReader sdr = com.ExecuteReader();

                while (sdr.Read() == true)
                {
                    output = (string)sdr["CommonValue2"];
                }

            }
            finally
            {
                sqlconn.Close();
            }


            return output;
        }

        public void InsertErrorTable(int ID, string ProgName, string FileName, int i, string ErrorMsg, string tokuchobango, int count, int shubetu)
        {
            SqlConnection sqlconn = new SqlConnection(connStr);
            sqlconn.Open();
            var cmd = sqlconn.CreateCommand();
            SqlTransaction transaction = sqlconn.BeginTransaction();
            cmd.Transaction = transaction;

            try
            {
                cmd.CommandText = "INSERT INTO T_FileReadError(" +
                "FileReadErrorID, FileReadErrorLineNo,FileReadErrorDateTime,FileReadErrorPgmname," +
                "FileReadErrorFilename,FileReadErrorMessage,FileReadErrorTokuchoBangou,FileReadErrorReadCount,FileReadErrorShubetsu) " +
                    " VALUES ( " + ID +
                " ," + i + " ,SYSDATETIME(),'" + ProgName + "',N'" + FileName + "', N'" +
                       ErrorMsg + "',N'" + tokuchobango + "'," + count + " ," + shubetu + ")";

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

        public int getReadCount(string TokuchoBango)
        {
            //同じファイル名で何度目のエラーか取得する
            int readCount = 1;
            using (var conn = new SqlConnection(connStr))
            {
                //エラーメッセージ
                conn.Open();
                var cmd = conn.CreateCommand();
                var dt = new DataTable();
                //SQL生成
                cmd.CommandText = "SELECT TOP 1IsNull(FileReadErrorReadCount,0) + 1 AS readCount " +
                "FROM T_FileReadError " +
                "WHERE FileReadErrorTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + TokuchoBango + "' " +
                "ORDER BY FileReadErrorReadCount DESC";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
                if (dt.Rows.Count >= 1)
                {
                    DataRow dr = dt.Rows[0];
                    readCount = int.Parse(dr["readCount"].ToString());
                }

                conn.Close();
            }

            return readCount;
        }
        public int getSaiban(string SaibanName)
        {
            //同じファイル名で何度目のエラーか取得する
            int saibanPlanNo = 0;
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                //採番テーブル取得
                var dt = new DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                  "SaibanNo + SaibanCountupNo AS SaibanNo " +
                  "FROM " + "M_Saiban " +
                  "WHERE SaibanMei = '" + SaibanName + "' ";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);

                if (dt.Rows.Count >= 1)
                {
                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    try
                    {
                        saibanPlanNo = int.Parse(dt.Rows[0][0].ToString());
                        //採番No（SaibanNo）を更新
                        cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                            saibanPlanNo + " WHERE SaibanMei = '" + SaibanName + "' ";
                        cmd.ExecuteNonQuery();

                        transaction.Commit();
                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                    }
                }
                else
                {
                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    try
                    {
                        cmd.CommandText = "INSERT INTO M_Saiban ( SaibanMei,SaibanNo,SaibanCountupNo,SaibanStartNo, SaibanMaxNo) VALUES( '" + SaibanName + "' ,1,1,1,0)";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        transaction.Commit();

                        saibanPlanNo = 1;
                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                    }
                }
                conn.Close();
            }

            return saibanPlanNo;
        }

        public void Insert_History(string KojinCD, string KojinNM, string BushoCD, string BushoNM, string Naiyo, string ProgramNM, string MadoguchiID)
        {
            SqlConnection sqlconn = new SqlConnection(connStr);
            sqlconn.Open();
            var cmd = sqlconn.CreateCommand();
            SqlTransaction transaction = sqlconn.BeginTransaction();
            cmd.Transaction = transaction;
            string madoguchiID = "null";
            string tokuchoBangou = "";

            if (MadoguchiID != "")
            {
                madoguchiID = "'" + MadoguchiID + "'";

                // 特調番号を取得する
                DataTable dt = new DataTable();
                dt = getData("MadoguchiUketsukeBangouEdaban", "MadoguchiUketsukeBangou", "MadoguchiJouhou with(nolock) ", "MadoguchiID = '" + MadoguchiID + "' ");

                if (dt != null && dt.Rows.Count > 0)
                {
                    tokuchoBangou = dt.Rows[0][0].ToString() + "-" + dt.Rows[0][1].ToString();
                }
            }
            try
            {
                int Histry_No = getSaiban("HistoryID");
                cmd.CommandText = "INSERT INTO T_HISTORY(" +
                "H_DATE_KEY, H_NO_KEY,H_OPERATE_DT," +
                "H_OPERATE_USER_ID,H_OPERATE_USER_MEI,H_OPERATE_USER_BUSHO_CD,H_OPERATE_USER_BUSHO_MEI, " +
                "H_OPERATE_NAIYO,H_ProgramName,MadoguchiID,HistoryBeforeTantoubushoCD,HistoryBeforeTantoushaCD,HistoryAfterTantoubushoCD,HistoryAfterTantoushaCD,H_TOKUCHOBANGOU) " +
                " VALUES ( " +
                   //" SYSDATETIME() ," + Histry_No + ",'" + DateTime.Today + "' " +
                   " SYSDATETIME() ," + Histry_No + ",'" + DateTime.Now + "' " +
                    " ,'" + KojinCD + "' ,N'" + KojinNM + "' ,'" + BushoCD + "' ,N'" + BushoNM + "' " +
                    " ,N'" + Naiyo + "' ,'" + ProgramNM + "' ," + madoguchiID + " ,NULL ,NULL ,NULL ,NULL,N'" + tokuchoBangou + "')";
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

        // 窓口ミハルの調査品目取込
        public string[] InsertHinmoku(string filePath, string MadoguchiID, string shoriUser, string shoriUserBusho, string flg = "0")
        {
            //Processオブジェクトを作成
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            //ComSpec(cmd.exe)のパスを取得して、FileNameプロパティに指定
            p.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec");

            //出力を読み取れるようにする
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardInput = false;
            //ウィンドウを表示しないようにする
            p.StartInfo.CreateNoWindow = true;
            // GeneXusのexeは共有フォルダに配置する
            // parm(in:&p_Filename,in:&P_MadoguchiID,out:&p_Ret);
            //　　　　ファイルパス    窓口ID の並びでパラメータを渡す

            // flg 0:調査品目取込 1:窓口一括
            // ファイルパスに半角スペースがあるとパラメータが分割出来ないので、"で囲む
            filePath = '"' + filePath + '"';
            p.StartInfo.Arguments = @"/c " + GetCommonValue1("CHOUSAHINMOKU_EXE_FOLDER") + " " + filePath + " " + MadoguchiID + " " + shoriUser + " " + shoriUserBusho + " " + flg;

            //起動
            p.Start();

            //出力を読み取る
            string results = p.StandardOutput.ReadToEnd();

            //プロセス終了まで待機する
            //WaitForExitはReadToEndの後である必要がある
            //(親プロセス、子プロセスでブロック防止のため)
            p.WaitForExit();
            p.Close();

            string[] result = results.Replace(Environment.NewLine, "").Split('|');

            return result;
        }

        // 自分大臣の単価取込
        public string[] InsertTanka(string filePath, string MadoguchiID, int i_ZumenNo, int i_Hinmei, int i_hachuu, string HinmokuChousainCD, string i_ShuFuku, string i_FileName, string shoriUser, string shoriUserBusho)
        {
            //Processオブジェクトを作成
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            //ComSpec(cmd.exe)のパスを取得して、FileNameプロパティに指定
            p.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec");

            //出力を読み取れるようにする
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardInput = false;
            //ウィンドウを表示しないようにする
            p.StartInfo.CreateNoWindow = true;
            // GeneXusのexeは共有フォルダに配置する
            // parm(in:&p_Filename,in:&P_MadoguchiID,out:&p_Ret);
            //　　　　ファイルパス    窓口ID の並びでパラメータを渡す
            // ファイルパスに全角スペースが含まれるので、'で囲む
            // ファイルパスに全角スペースが含まれるので、"で囲む
            filePath = '"' + filePath + '"';
            i_FileName = '"' + i_FileName + '"';
            p.StartInfo.Arguments = @"/c " + GetCommonValue1("TANKA_EXE_FOLDER") + " " + filePath + " " + MadoguchiID + " " + i_ZumenNo + " " + i_Hinmei + " " + i_hachuu + " " + HinmokuChousainCD + " " + i_ShuFuku + " " + i_FileName + " " + " " + shoriUser + " " + shoriUserBusho;

            //起動
            p.Start();

            //出力を読み取る
            string results = p.StandardOutput.ReadToEnd();

            //プロセス終了まで待機する
            //WaitForExitはReadToEndの後である必要がある
            //(親プロセス、子プロセスでブロック防止のため)
            p.WaitForExit();
            p.Close();

            string[] result = results.Replace(Environment.NewLine, "").Split('|');

            return result;
        }

        public string[] InsertReportWork(int ListID, string UserID, string[] data)
        {
            SqlConnection sqlconn = new SqlConnection(connStr);
            sqlconn.Open();
            var cmd = sqlconn.CreateCommand();
            SqlTransaction transaction = sqlconn.BeginTransaction();
            cmd.Transaction = transaction;

            int WorkID = getSaiban("ReportWorkID");
            try
            {
                cmd.CommandText = "INSERT INTO T_ReportWork(" +
                "ReportWorkID, ReportWorkUserID, ReportWorkDateTime, ReportWorkPrintListID " +
                ") VALUES ( " +
                WorkID + ",'" + UserID + "', SYSDATETIME() " + "," + ListID + " )";

                cmd.ExecuteNonQuery();

                // ListID 2:エントリーチェックシート出力
                // えんとり君修正STEP2 赤黒まとめ出力
                //if (ListID == 1 || ListID == 2)
                if (ListID == 1 || ListID == 2 || ListID == 352 || ListID == 353)
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES ( " +
                    WorkID + ",1,'AnkenJouhouID', 2 ,null, " + data[0] + " ,null ) ";

                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES ( " +
                    WorkID + ",2,'AnkenJutakuBangou', 1 , N'" + data[1] + "' ,null, null ) ";

                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES ( " +
                    WorkID + ",3,'kakuninflg', 2 ,null," + data[2] + ", null ) ";

                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES ( " +
                    WorkID + ",4,'PrintShurui', 2 ,null," + data[3] + ", null ) ";

                    cmd.ExecuteNonQuery();
                }

                // ListID 44:調査品目一覧
                if (ListID == 44)
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'MadoguchiID'" + ",2" + ",null" + "," + ChangeSqlText(data[0], 0, 0) + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'Shozoku'" + ",1" + ",N'" + ChangeSqlText(data[1], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",3" + ",'HinmokuChousain'" + ",1" + ",N'" + ChangeSqlText(data[2], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",4" + ",'ShuFuku'" + ",2" + ",null" + "," + ChangeSqlText(data[3], 0, 0) + ",null" + ")," +
                    "(" + WorkID + ",5" + ",'ChousaHinmei'" + ",1" + ",N'" + ChangeSqlText(data[4], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",6" + ",'ChousaKikaku'" + ",1" + ",N'" + ChangeSqlText(data[5], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",7" + ",'Zaikou'" + ",2" + ",null" + "," + ChangeSqlText(data[6], 0, 0) + ",null" + ")," +
                    "(" + WorkID + ",8" + ",'TantoushaKuuhaku'" + ",2" + ",null" + "," + ChangeSqlText(data[7], 0, 0) + ",null" + ")," +
                    "(" + WorkID + ",9" + ",'PrintGamen'" + ",2" + ",null" + "," + data[8] + ",null" + ")";

                    cmd.ExecuteNonQuery();

                }
                // 集計表
                if (ListID == 47)
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'MadoguchiID'" + ",2" + ",null" + "," + data[0] + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'ZeninSyukeihyo'" + ",2" + ",null" + "," + data[1] + ",null" + ")," +
                    "(" + WorkID + ",3" + ",'ShibuMei'" + ",1" + ",N'" + ChangeSqlText(data[2], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",4" + ",'KojinCD'" + ",2" + ",null" + "," + data[3] + ",null" + ")," +
                    "(" + WorkID + ",5" + ",'ShuFuku'" + ",2" + ",null" + "," + data[4] + ",null" + ")," +
                    "(" + WorkID + ",6" + ",'FileName'" + ",1" + ",N'" + ChangeSqlText(data[5], 0, 0) + "'" + ",null" + ",null)";

                    cmd.ExecuteNonQuery();
                }
                // ListID 230:エントリくん一覧出力(新）
                // えんとり君修正STEP3 「1002:エントリくん一覧出力(フラット版）」と「1003:経理向け帳票」追加
                if (ListID == 230 || ListID == 1002 || ListID == 1003)
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'UriageNendo'" + ",1" + ",N'" + ChangeSqlText(data[0], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'UriageNendoOption'" + ",2" + ",null" + "," + data[1] + "" + ",null" + ")," +
                    "(" + WorkID + ",3" + ",'JutakuKashoshibuCD'" + ",1" + ",N'" + ChangeSqlText(data[2], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",4" + ",'JigyoubuCD'" + ",1" + ",N'" + ChangeSqlText(data[3], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",5" + ",'AnkenKubunCD'" + ",1" + ",N'" + ChangeSqlText(data[4], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",6" + ",'NyuusatsuYoteibiFrom'" + ",3" + ",null" + ",null" + "," + data[5] + ")," +
                    "(" + WorkID + ",7" + ",'NyuusatsuYoteibiTo'" + ",3" + ",null" + ",null" + "," + data[6] + ")," +
                    "(" + WorkID + ",8" + ",'KeiyakuKubunCD'" + ",2" + ",null" + "," + data[7] + ",null" + ")," +
                    "(" + WorkID + ",9" + ",'NyusatsuJokyouCD'" + ",2" + ",null" + "," + data[8] + "" + ",null" + ")," +
                    "(" + WorkID + ",10" + ",'Rakusatsusha'" + ",1" + ",N'" + ChangeSqlText(data[9], 0, 0) + "'" + ",null" + ",null" + ")," +
                    //不具合No1357（1128）nullで渡していたが、値を渡すよう修正。呼び出し元ではdata[10]にはしっかり値を入れていた。
                    "(" + WorkID + ",11" + ",'HachushaKubun1'" + ",2" + ",null" + "," + data[10] + ",null" + ")," +
                    //"(" + WorkID + ",11" + ",'HachushaKubun1'" + ",2" + ",null" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",12" + ",'KeikakuBangou'" + ",1" + ",N'" + ChangeSqlText(data[11], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",13" + ",'KeikakuAnkenMei'" + ",1" + ",N'" + ChangeSqlText(data[12], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",14" + ",'JutakuBangou'" + ",1" + ",N'" + ChangeSqlText(data[13], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",15" + ",'AnkenBangou'" + ",1" + ",N'" + ChangeSqlText(data[14], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",16" + ",'GyoumuMei'" + ",1" + ",N'" + ChangeSqlText(data[15], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",17" + ",'HachuushaKaMei'" + ",1" + ",N'" + ChangeSqlText(data[16], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",18" + ",'Kanriduki'" + ",3" + ",null" + ",null" + "," + data[17] + ")," +
                    "(" + WorkID + ",19" + ",'SankouMitsumori'" + ",2" + ",null" + "," + data[18] + "" + ",null" + ")," +
                    "(" + WorkID + ",20" + ",'JyutyuIyoku'" + ",2" + ",null" + "," + data[19] + "" + ",null" + ")," +
                    "(" + WorkID + ",21" + ",'Hikiaijhokyo'" + ",2" + ",null" + "," + data[20] + "" + ",null" + ")," +
                    "(" + WorkID + ",22" + ",'ToukaiOusatu'" + ",2" + ",null" + "," + data[21] + "" + ",null" + ")," +
                    "(" + WorkID + ",23" + ",'NendogoeHaibun'" + ",2" + ",null" + "," + data[22] + "" + ",null" + ")," +
                    "(" + WorkID + ",24" + ",'HachuushaCD'" + ",1" + ",N'" + ChangeSqlText(data[23], 0, 0) + "'" + ",null" + "" + ",null" + ")," +
                    "(" + WorkID + ",25" + ",'KianJokyo'" + ",1" + ",N'" + ChangeSqlText(data[24], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",26" + ",'SaishinDenpyou'" + ",2" + ",null" + "," + data[25] + "" + ",null" + ")," +
                    "(" + WorkID + ",27" + ",'HyouziKensuu'" + ",2" + ",null" + "," + data[26] + "" + ",null" + ")," +
                    "(" + WorkID + ",28" + ",'KoukiKaishiNendo'" + ",1" + ",N'" + ChangeSqlText(data[27], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",29" + ",'KoukiKaishiNendoOption'" + ",2" + ",null" + "," + data[28] + "" + ",null" + ")," +
                    //えんとり君修正STEP2　案件情報一覧検索画面のエントリくん一覧帳票のパラメータ追加
                    "(" + WorkID + ",30" + ",'Ribicyou'" + ",1" + ",null" + "," + data[29] + "" + ",null" + ")," +
                    "(" + WorkID + ",31" + ",'SashaKeiyu'" + ",1" + ",null" + "," + data[30] + "" + ",null" + ")," +
                    //AnkenBangou2パラメータ追加
                    "(" + WorkID + ",32" + ",'AnkenBangou2'" + ",1" + ",N'" + ChangeSqlText(data[31], 0, 0) + "'" + ",null" + ",null)," +
                    //えんとり君修正STEP2　案件情報一覧検索画面のエントリくん一覧帳票のパラメータ追加
                    "(" + WorkID + ",33" + ",'AnkenJizenDashinCheck'" + ",1" + ",null" + "," + data[32] + "" + ",null" + ")," +
                    "(" + WorkID + ",34" + ",'AnkenNyuusatuCheck'" + ",1" + ",null" + "," + data[33] + "" + ",null" + ")," +
                    "(" + WorkID + ",35" + ",'AnkenKeiyakuCheck'" + ",1" + ",null" + "," + data[34] + "" + ",null" + ")"
                    ;

                    cmd.ExecuteNonQuery();

                }
                // ListID 3:契約図書保管チェック表、21:単価契約見積書書式集、22:着手完了届書式集、23:使用印鑑簿、52:ISO書式集
                if (ListID == 3 || ListID == 21 || ListID == 22 || ListID == 23 || ListID == 52)
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'AnkenJouhouID'" + ",2" + ",null" + "," + data[0] + ",null" + ")";

                    cmd.ExecuteNonQuery();
                }
                // ListID 42:受託実績統括表
                if (ListID == 42)
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'NendoID'" + ",2" + ",null" + "," + data[0] + ",null" + ")";

                    cmd.ExecuteNonQuery();
                }
                transaction.Commit();
            }
            catch
            {
                transaction.Rollback();
                return null;
            }
            finally
            {
                sqlconn.Close();
            }

            //Processオブジェクトを作成
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            //ComSpec(cmd.exe)のパスを取得して、FileNameプロパティに指定
            p.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec");

            //出力を読み取れるようにする
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardInput = false;
            //ウィンドウを表示しないようにする
            p.StartInfo.CreateNoWindow = true;
            //コマンドラインを指定（"/c"は実行後閉じるために必要）
            //p.StartInfo.Arguments = @"/c " + System.Environment.CurrentDirectory + @"\Resource\module\aproexceloutreportmain.exe" + " " + WorkID;

            // 47:集計表
            if (ListID == 47)
            {
                // 集計表は呼ぶPCが違う
                // GeneXusのexeは共有フォルダに配置する aproexceloutreportmadoguchi.exe
                p.StartInfo.Arguments = @"/c " + GetCommonValue1("MADOGUCHI_EXE_FOLDER") + " " + WorkID;
            }
            // 調査品目明細一覧出力
            else if (ListID == 44)
            {
                // GeneXusのexeは共有フォルダに配置する aproexceloutreportmadoguchi.exe
                p.StartInfo.Arguments = @"/c " + GetCommonValue1("MADOGUCHI_EXE_FOLDER") + " " + WorkID;
            }
            else
            {
                // GeneXusのexeは共有フォルダに配置する aproexceloutreportmain.exe
                p.StartInfo.Arguments = @"/c " + GetCommonValue1("REPORT_EXE_FOLDER") + " " + WorkID;
            }

            //if (ListID != 47) 
            //{
            //    // GeneXusのexeは共有フォルダに配置する aproexceloutreportmain.exe
            //    p.StartInfo.Arguments = @"/c " + GetCommonValue1("REPORT_EXE_FOLDER") + " " + WorkID;
            //}
            //else if(ListID == 47)
            //{
            //    // 集計表は呼ぶPCが違う
            //    // GeneXusのexeは共有フォルダに配置する aproexceloutreportmadoguchi.exe
            //    p.StartInfo.Arguments = @"/c " + GetCommonValue1("MADOGUCHI_EXE_FOLDER") + " " + WorkID;
            //}

            //起動
            p.Start();

            //出力を読み取る
            string results = p.StandardOutput.ReadToEnd();

            //プロセス終了まで待機する
            //WaitForExitはReadToEndの後である必要がある
            //(親プロセス、子プロセスでブロック防止のため)
            p.WaitForExit();
            p.Close();

            string[] result = results.Replace(Environment.NewLine, "").Split('|');

            return result;
        }

        //public string[] InsertMadoguchiReportWork(int ListID, string UserID, string[] data, string reportType)
        public string[] InsertMadoguchiReportWork(int ListID, string UserID, string[] data, string reportType, string printDataPattern = "")
        {
            SqlConnection sqlconn = new SqlConnection(connStr);
            sqlconn.Open();
            var cmd = sqlconn.CreateCommand();
            SqlTransaction transaction = sqlconn.BeginTransaction();
            cmd.Transaction = transaction;

            int WorkID = getSaiban("ReportWorkID");
            try
            {
                cmd.CommandText = "INSERT INTO T_ReportWork(" +
                "ReportWorkID, ReportWorkUserID, ReportWorkDateTime, ReportWorkPrintListID " +
                ") VALUES ( " +
                WorkID + ",'" + UserID + "', SYSDATETIME() " + "," + ListID + " )";

                cmd.ExecuteNonQuery();
                // 集計表
                if ("Shukeihyo".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'MadoguchiID'" + ",2" + ",null" + "," + data[0] + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'ZeninSyukeihyo'" + ",2" + ",null" + "," + data[1] + ",null" + ")," +
                    "(" + WorkID + ",3" + ",'ShibuMei'" + ",1" + ",N'" + ChangeSqlText(data[2], 0, 0) + "'" + ",null" + ",null" + ")," +
                    "(" + WorkID + ",4" + ",'KojinCD'" + ",2" + ",null" + "," + data[3] + ",null" + ")," +
                    "(" + WorkID + ",5" + ",'ShuFuku'" + ",2" + ",null" + "," + data[4] + ",null" + ")," +
                    "(" + WorkID + ",6" + ",'FileName'" + ",1" + ",N'" + ChangeSqlText(data[5], 0, 0) + "'" + ",null" + ",null)," +
                    "(" + WorkID + ",7" + ",'PrintGamen'" + ",2" + ",null" + "," + data[6] + ",null" + ")";
                }

                // 協力依頼書
                if ("KyouryokuIrai".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'MadoguchiID'" + ",2" + ",null" + "," + data[0] + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'YobidashiMoto'" + ",2" + ",null" + "," + data[1] + ",null" + ")";
                }

                // 業務連絡票
                if ("Gyoumurenraku".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'MadoguchiID'" + ",2" + ",null" + "," + data[0] + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'GyoumuBushoCD'" + ",1" + ",N'" + data[1] + "',null" + ",null" + ")";
                }

                // 協力依頼結果送付書
                if ("Kekkasouhusho".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'MadoguchiID'" + ",2" + ",null" + "," + data[0] + ",null" + ")";
                }

                // 報告書または担当者一覧
                if ("Houkokusho".Equals(reportType) || "TantoushaIchiran".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'MadoguchiID'" + ",2" + ",null" + "," + data[0] + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'PrintGamen'" + ",2" + ",null" + "," + data[1] + ",null" + ")," +
                    "(" + WorkID + ",3" + ",'KikanStart'" + ",3" + ",null" + ",null" + "," + data[2] + ")," +
                    "(" + WorkID + ",4" + ",'KikanEnd'" + ",3" + ",null" + ",null" + "," + data[3] + ")," +
                    "(" + WorkID + ",5" + ",'BushoCD'" + ",1" + ", N'" + ChangeSqlText(data[4], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",6" + ",'Tantousha'" + ",1" + ", N'" + ChangeSqlText(data[5], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",7" + ",'HoukokuSentaku'" + ",2" + ",null" + "," + data[6] + ",null" + ")," +
                    "(" + WorkID + ",8" + ",'seikyuuGetsu'" + ",1" + ", N'" + ChangeSqlText(data[7], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",9" + ",'ShuFuku'" + ",2" + ",null" + "," + data[8] + ",null" + ")," +
                    "(" + WorkID + ",10" + ",'Hinmei'" + ",1" + ", N'" + ChangeSqlText(data[9], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",11" + ",'Kikaku'" + ",1" + ", N'" + ChangeSqlText(data[10], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",12" + ",'Zaikou'" + ",2" + ",null" + "," + data[11] + ",null" + ")," +
                    "(" + WorkID + ",13" + ",'KuhakuList'" + ",2" + ",null" + "," + data[12] + ",null" + ")," +
                    "(" + WorkID + ",14" + ",'Shimekiribi'" + ",3" + ",null" + ",null" + "," + data[13] + ")," +
                    "(" + WorkID + ",15" + ",'HizukeKubun'" + ",2" + ",null" + "," + data[14] + ",null" + ")," +
                    "(" + WorkID + ",16" + ",'Memo1'" + ",1" + ", N'" + ChangeSqlText(data[15], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",17" + ",'Memo2'" + ",1" + ", N'" + ChangeSqlText(data[16], 0, 0) + "',null" + ",null" + ")";

                    // えんとり君修正STEP2 報告書共通化
                    if (printDataPattern.Equals("800") || printDataPattern.Equals("801"))
                    {
                        cmd.CommandText = cmd.CommandText + ",(" + WorkID + ",18" + ",'ChuushiYouhi'" + ",2" + ",null" + "," + data[17] + ",null" + ")";
                    }
                }

                // 工程表
                if ("KouteiHyo".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'PrintTokumei'" + ",2" + ",null" + "," + data[0] + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'BushoMei'" + ",1" + ", N'" + ChangeSqlText(data[1], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",3" + ",'ChousainMei'" + ",1" + ", N'" + ChangeSqlText(data[2], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",4" + ",'NendoID'" + ",2" + ",null" + "," + data[3] + ",null" + ")," +
                    "(" + WorkID + ",5" + ",'NendoOption'" + ",2" + ",null" + "," + data[4] + ",null" + ")," +
                    "(" + WorkID + ",6" + ",'ChousaBushoCD'" + ",1" + ", N'" + ChangeSqlText(data[5], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",7" + ",'HachuuKikanmei'" + ",1" + ", N'" + ChangeSqlText(data[6], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",8" + ",'UketsukeBangou'" + ",1" + ", N'" + ChangeSqlText(data[7], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",9" + ",'Shimekiribi'" + ",3" + ",null" + ",null" + "," + data[8] + ")," +
                    "(" + WorkID + ",10" + ",'Shimekiribi_e'" + ",3" + ",null" + ",null" + "," + data[9] + ")," +
                    "(" + WorkID + ",11" + ",'ShinchokuJoukyou'" + ",2" + ",null" + "," + data[10] + ",null" + ")," +
                    "(" + WorkID + ",12" + ",'MadoguchiBushoCD'" + ",1" + ", N'" + ChangeSqlText(data[11], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",13" + ",'GyoumuMeishou'" + ",1" + ", N'" + ChangeSqlText(data[12], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",14" + ",'KanriBangou'" + ",1" + ", N'" + ChangeSqlText(data[13], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",15" + ",'KikanShitei'" + ",2" + ",null" + "," + data[14] + ",null" + ")," +
                    "(" + WorkID + ",16" + ",'ChousaShubetsu'" + ",2" + ",null" + "," + data[15] + ",null" + ")," +
                    "(" + WorkID + ",17" + ",'MadoguchiTantousha'" + ",1" + ", N'" + ChangeSqlText(data[16], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",18" + ",'KoujiKenmei'" + ",1" + ", N'" + ChangeSqlText(data[17], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",19" + ",'ShimekiriShitei'" + ",2" + ",null" + "," + data[18] + ",null" + ")," +
                    "(" + WorkID + ",20" + ",'ChousaKubunJibusho'" + ",2" + ",null" + "," + data[19] + ",null" + ")," +
                    "(" + WorkID + ",21" + ",'ChousaKubunShibuShibu'" + ",2" + ",null" + "," + data[20] + ",null" + ")," +
                    "(" + WorkID + ",22" + ",'ChousaKubunHonbuShibu'" + ",2" + ",null" + "," + data[21] + ",null" + ")," +
                    "(" + WorkID + ",23" + ",'ChousaKubunShibuHonbu'" + ",2" + ",null" + "," + data[22] + ",null" + ")," +
                    "(" + WorkID + ",24" + ",'ChousaHinmoku'" + ",1" + ", N'" + ChangeSqlText(data[23], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",25" + ",'Shinchoku'" + ",2" + ",null" + "," + data[24] + ",null" + ")," +
                    "(" + WorkID + ",26" + ",'HonbuTanpin'" + ",2" + ",null" + "," + data[25] + ",null" + ")," +
                    "(" + WorkID + ",27" + ",'ChousaTantousha'" + ",1" + ", N'" + ChangeSqlText(data[26], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",28" + ",'Memo'" + ",1" + ", N'" + ChangeSqlText(data[27], 0, 0) + "',null" + ",null" + ")";
                }

                // 調査状況一覧
                if ("ChousaJokyou".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'BushoCD'" + ",1" + ", N'" + ChangeSqlText(data[0], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'NendoID'" + ",2" + ",null" + "," + data[1] + ",null" + ")," +
                    "(" + WorkID + ",3" + ",'NendoOption'" + ",2" + ",null" + "," + data[2] + ",null" + ")," +
                    "(" + WorkID + ",4" + ",'ChousaBushoCD'" + ",1" + ", N'" + ChangeSqlText(data[3], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",5" + ",'HachuuKikanmei'" + ",1" + ", N'" + ChangeSqlText(data[4], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",6" + ",'UketsukeBangou'" + ",1" + ", N'" + ChangeSqlText(data[5], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",7" + ",'Shimekiribi'" + ",3" + ",null" + ",null" + "," + data[6] + ")," +
                    "(" + WorkID + ",8" + ",'Shimekiribi_e'" + ",3" + ",null" + ",null" + "," + data[7] + ")," +
                    "(" + WorkID + ",9" + ",'ShinchokuJoukyou'" + ",2" + ",null" + "," + data[8] + ",null" + ")," +
                    "(" + WorkID + ",10" + ",'MadoguchiBushoCD'" + ",1" + ", N'" + ChangeSqlText(data[9], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",11" + ",'GyoumuMeishou'" + ",1" + ", N'" + ChangeSqlText(data[10], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",12" + ",'KanriBangou'" + ",1" + ", N'" + ChangeSqlText(data[11], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",13" + ",'KoujiKenmei'" + ",1" + ", N'" + ChangeSqlText(data[12], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",14" + ",'ChousaKubunJibusho'" + ",2" + ",null" + "," + data[13] + ",null" + ")," +
                    "(" + WorkID + ",15" + ",'ChousaKubunShibuShibu'" + ",2" + ",null" + "," + data[14] + ",null" + ")," +
                    "(" + WorkID + ",16" + ",'ChousaKubunHonbuShibu'" + ",2" + ",null" + "," + data[15] + ",null" + ")," +
                    "(" + WorkID + ",17" + ",'ChousaKubunShibuHonbu'" + ",2" + ",null" + "," + data[16] + ",null" + ")," +
                    "(" + WorkID + ",18" + ",'Shinchoku'" + ",2" + ",null" + "," + data[17] + ",null" + ")," +
                    "(" + WorkID + ",19" + ",'HonbuTanpin'" + ",2" + ",null" + "," + data[18] + ",null" + ")," +
                    "(" + WorkID + ",20" + ",'ChousaTantousha'" + ",1" + ", N'" + ChangeSqlText(data[19], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",21" + ",'Memo'" + ",1" + ", N'" + ChangeSqlText(data[20], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",22" + ",'ShibuBikou'" + ",1" + ", N'" + ChangeSqlText(data[21], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",23" + ",'ShuFuku'" + ",2" + ",null" + "," + data[22] + ",null" + ")";
                }

                // ISO書式集
                if ("ISOShosiki".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'MadoguchiID'" + ",2" + ",null" + "," + data[0] + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'PrintGamen'" + ",2" + ",null" + "," + data[1] + ",null" + ")";
                }

                // ランク内訳明細
                if ("RankUchiwakeMeisai".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'MadoguchiID'" + ",2" + ",null" + "," + data[0] + ",null" + ")";
                }

                // 窓口ミハル一覧 または 管理帳票
                if ("MiharuIchiran".Equals(reportType) || "KanriChouhyou".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'NendoID'" + ",2" + ",null" + "," + data[0] + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'NendoOption'" + ",2" + ",null" + "," + data[1] + ",null" + ")," +
                    "(" + WorkID + ",3" + ",'ChousaBushoCD'" + ",1" + ", N'" + ChangeSqlText(data[2], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",4" + ",'HonbuTanpin'" + ",2" + ",null" + "," + data[3] + ",null" + ")," +
                    "(" + WorkID + ",5" + ",'HachuuKikanmei'" + ",1" + ", N'" + ChangeSqlText(data[4], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",6" + ",'UketsukeBangou'" + ",1" + ", N'" + ChangeSqlText(data[5], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",7" + ",'DateSelect'" + ",2" + ",null" + "," + data[6] + ",null" + ")," +
                    "(" + WorkID + ",8" + ",'Shimekiribi'" + ",3" + ",null" + ",null" + "," + data[7] + ")," +
                    "(" + WorkID + ",9" + ",'Shimekiribi_e'" + ",3" + ",null" + ",null" + "," + data[8] + ")," +
                    "(" + WorkID + ",10" + ",'JutakuBusho'" + ",1" + ", N'" + ChangeSqlText(data[9], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",11" + ",'JutakuBushoTantousha'" + ",1" + ", N'" + ChangeSqlText(data[10], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",12" + ",'MadoguchiBushoCD'" + ",1" + ", N'" + ChangeSqlText(data[11], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",13" + ",'MadoguchiBushoTantousha'" + ",1" + ", N'" + ChangeSqlText(data[12], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",14" + ",'ChousaKubunJibusho'" + ",2" + ",null" + "," + data[13] + ",null" + ")," +
                    "(" + WorkID + ",15" + ",'ChousaKubunShibuShibu'" + ",2" + ",null" + "," + data[14] + ",null" + ")," +
                    "(" + WorkID + ",16" + ",'ChousaKubunHonbuShibu'" + ",2" + ",null" + "," + data[15] + ",null" + ")," +
                    "(" + WorkID + ",17" + ",'ChousaKubunShibuHonbu'" + ",2" + ",null" + "," + data[16] + ",null" + ")," +
                    "(" + WorkID + ",18" + ",'GyoumuMeishou'" + ",1" + ", N'" + ChangeSqlText(data[17], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",19" + ",'KanriBangou'" + ",1" + ", N'" + ChangeSqlText(data[18], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",20" + ",'KoujiKenmei'" + ",1" + ", N'" + ChangeSqlText(data[19], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",21" + ",'dashinnchuu'" + ",2" + ",null" + "," + data[20] + ",null" + ")," +
                    "(" + WorkID + ",22" + ",'KikanShitei'" + ",2" + ",null" + "," + data[21] + ",null" + ")," +
                    "(" + WorkID + ",23" + ",'ShimekiriShitei'" + ",2" + ",null" + "," + data[22] + ",null" + ")," +
                    "(" + WorkID + ",24" + ",'ChousaShubetsu'" + ",2" + ",null" + "," + data[23] + ",null" + ")," +
                    "(" + WorkID + ",25" + ",'Shijisho'" + ",2" + ",null" + "," + data[24] + ",null" + ")," +
                    "(" + WorkID + ",26" + ",'JiishiKubun'" + ",2" + ",null" + "," + data[25] + ",null" + ")," +
                    "(" + WorkID + ",27" + ",'ChousaHinmoku'" + ",1" + ", N'" + ChangeSqlText(data[26], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",28" + ",'Kanryou_sita'" + ",2" + ",null" + "," + data[27] + ",null" + ")," +
                    "(" + WorkID + ",29" + ",'ShinchokuJoukyou'" + ",2" + ",null" + "," + data[28] + ",null" + ")";
                }

                // 業務完了内訳表
                if ("GyoumuKanryou".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'TankaKeiyakuID'" + ",2" + ",null" + "," + data[0] + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'HoukokuSentaku'" + ",2" + ",null" + "," + data[1] + ",null" + ")," +
                    "(" + WorkID + ",3" + ",'KikanStart'" + ",3" + ",null" + ",null" + "," + data[2] + ")," +
                    "(" + WorkID + ",4" + ",'KikanEnd'" + ",3" + ",null" + ",null" + "," + data[3] + ")," +
                    "(" + WorkID + ",5" + ",'seikyuuGetsu'" + ",1" + ", N'" + ChangeSqlText(data[4], 0, 0) + "',null" + ",null" + ")";
                }

                // 部所別提出状況一覧表
                if ("BushoBetsu".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'MadoguchiID'" + ",2" + ",null" + "," + data[0] + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'BushoCD'" + ",1" + ", N'" + ChangeSqlText(data[1], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",3" + ",'PrintGamen'" + ",2" + ",null" + "," + data[2] + ",null" + ")";
                }

                // 窓口ミハル一括取込用
                if ("IkkatsuTorikomi".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'IkkatsuTorikomi'" + ",2" + ",null" + "," + data[0] + ",null" + ")";
                }

                // えんとり君修正STEP2：単価契約の報告書共通化
                if ("TankaHoukokusho".Equals(reportType))
                {
                    cmd.CommandText = "INSERT INTO T_ReportWorkDetail(" +
                    "ReportWorkID, ReportWorkDetailID, ReportWorkDetailColumn, ReportWorkDetailType, ReportWorkDetailStr, ReportWorkDetailInt, ReportWorkDetailDate " +
                    ") VALUES " +
                    "(" + WorkID + ",1" + ",'TankaKeiyakuID'" + ",2" + ",null" + "," + data[0] + ",null" + ")," +
                    "(" + WorkID + ",2" + ",'HoukokuSentaku'" + ",2" + ",null" + "," + data[1] + ",null" + ")," +
                    "(" + WorkID + ",3" + ",'KikanStart'" + ",3" + ",null" + ",null" + "," + data[2] + ")," +
                    "(" + WorkID + ",4" + ",'KikanEnd'" + ",3" + ",null" + ",null" + "," + data[3] + ")," +
                    "(" + WorkID + ",5" + ",'seikyuuGetsu'" + ",1" + ", N'" + ChangeSqlText(data[4], 0, 0) + "',null" + ",null" + ")," +
                    "(" + WorkID + ",6" + ",'HoukokuSentaku2'" + ",2" + ",null" + "," + data[5] + ",null" + ")," +
                    "(" + WorkID + ",7" + ",'ChuushiYouhi'" + ",2" + ",null" + "," + data[6] + ",null" + ")";
                }
                cmd.ExecuteNonQuery();
                transaction.Commit();
            }
            catch
            {
                transaction.Rollback();
                return null;
            }
            finally
            {
                sqlconn.Close();
            }

            //Processオブジェクトを作成
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            //ComSpec(cmd.exe)のパスを取得して、FileNameプロパティに指定
            p.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec");

            //出力を読み取れるようにする
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardInput = false;
            //ウィンドウを表示しないようにする
            p.StartInfo.CreateNoWindow = true;
            //コマンドラインを指定（"/c"は実行後閉じるために必要）
            //p.StartInfo.Arguments = @"/c " + System.Environment.CurrentDirectory + @"\Resource\module\aproexceloutreportmain.exe" + " " + WorkID;

            // 集計表は呼ぶPCが違う
            // GeneXusのexeは共有フォルダに配置する
            p.StartInfo.Arguments = @"/c " + GetCommonValue1("MADOGUCHI_EXE_FOLDER") + " " + WorkID;

            //起動
            p.Start();

            //出力を読み取る
            string results = p.StandardOutput.ReadToEnd();

            //プロセス終了まで待機する
            //WaitForExitはReadToEndの後である必要がある
            //(親プロセス、子プロセスでブロック防止のため)
            p.WaitForExit();
            p.Close();

            string[] result = results.Replace(Environment.NewLine, "").Split('|');

            return result;
        }

        // 別Window表示枚数制御
        public Boolean GetFormsShow(string System)
        {
            // return true:表示OK false：表示NG
            int num = 0;
            if ("Entry".Equals(System))
            {
                for (int i = 0; i < Application.OpenForms.Count; i++)
                {
                    Form f = Application.OpenForms[i];
                    if (f.Text.IndexOf("エントリくん") >= 0 && f.Visible)
                    {
                        num++;
                    }
                }
                if (num < 2)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else if ("Tokuchoyaro".Equals(System))
            {
                for (int i = 0; i < Application.OpenForms.Count; i++)
                {
                    Form f = Application.OpenForms[i];
                    if ((f.Text.IndexOf("特調野郎") >= 0 || f.Text.IndexOf("窓口") >= 0 || f.Text.IndexOf("自分大臣") >= 0 || f.Text.IndexOf("特命課長") >= 0) && f.Visible)
                    {
                        num++;
                    }
                }
                if (num < 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else if ("tanka".Equals(System))
            {
                for (int i = 0; i < Application.OpenForms.Count; i++)
                {
                    Form f = Application.OpenForms[i];
                    if ((f.Text.IndexOf("単価契約") >= 0 && f.Visible))
                    {
                        num++;
                    }
                }
                if (num < 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                // 上記以外なら開いてOK
                return true;
            }
        }

        // 不具合No1338 窓口ミハル更新SQL
        public Boolean MadoguchiUpdateRealTime_SQL(string MadoguchiID, string[,] data, out string mes, string[] UserInfos)
        {
            string methodName = ".MadoguchiUpdateRealTime_SQL";
            mes = "";
            SqlConnection sqlconn = new SqlConnection(connStr);
            sqlconn.Open();
            var cmd = sqlconn.CreateCommand();
            SqlTransaction transaction = sqlconn.BeginTransaction();
            cmd.Transaction = transaction;

            try
            {
                //窓口情報更新　報告済みはいずれにしても変更
                cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                    "MadoguchiHoukokuzumi = N'" + data[0, 34] + "' ";

                cmd.CommandText += " WHERE MadoguchiID =  " + MadoguchiID;
                cmd.ExecuteNonQuery();

                //報告済み
                if ("1".Equals(data[0, 34]))
                {
                    //実施区分により変更されているがここでは判定しなくてよいのか？
                    //その他情報も更新する。
                    //進捗状況 10：依頼　調査中→40：集計中　70：二次検済　80：中止

                    cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                                "MadoguchiL1ChousaShinchoku = 70 " +
                                ",MadoguchiL1ChousaKakunin = 1 " +
                                ",MadoguchiL1AsteriaKoushinFlag = 1 " +
                                ",MadoguchiL1UpdateDate = SYSDATETIME() " +
                                ",MadoguchiL1UpdateUser = N'" + UserInfos[0] + "' " +
                                ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
                                " WHERE MadoguchiL1ChousaShinchoku != 80 AND MadoguchiID = " + MadoguchiID;

                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                                "ChousaShinchokuJoukyou = 70 " +
                                ",ChousaHoukokuzumi = 1 " +
                                ",ChousaUpdateDate = SYSDATETIME() " +
                                ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +
                                ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                                " WHERE ChousaShinchokuJoukyou != 80 AND MadoguchiID = " + MadoguchiID;

                    cmd.ExecuteNonQuery();

                }
                transaction.Commit();

            }
            catch (ArithmeticException e)
            {
                transaction.Rollback();
                Console.WriteLine(e);
                return false;
            }
            finally
            {
                sqlconn.Close();
            }

            return true;
        }

        // 窓口ミハル更新SQL
        public Boolean MadoguchiUpdate_SQL(int tab, string MadoguchiID, string[,] data, out string mes, string[] UserInfos, string[,] data2 = null, string gamenMode = "")
        {
            string methodName = ".MadoguchiUpdate_SQL";
            mes = "";
            SqlConnection sqlconn = new SqlConnection(connStr);
            sqlconn.Open();
            var cmd = sqlconn.CreateCommand();
            SqlTransaction transaction = sqlconn.BeginTransaction();
            cmd.Transaction = transaction;

            try
            {
                //調査概要
                if (tab == 1)
                {
                    //M_COMMON_MASTERからCHOUSAKIJUNBIのデフォルトを取得する
                    var dtCommon = new DataTable();
                    cmd.CommandText = "SELECT CommonValue1 " +
                        "FROM M_CommonMaster " +
                        "WHERE CommonMasterKye = 'CHOUSAKIJUNBI_DEFAULT' ";
                    //データ取得
                    var sdaC = new SqlDataAdapter(cmd);
                    sdaC.Fill(dtCommon);

                    String CommonValue = "  年 月号　";
                    if (dtCommon.Rows.Count > 0)
                    {
                        CommonValue = dtCommon.Rows[0][0].ToString();
                    }

                    //新規登録の場合
                    if ("insert".Equals(data[0, 0]))
                    {
                        // 進捗状況
                        // 実施区分が中止→80:中止
                        // 実施区分が中止でなく、報告済み→70:二次検証済
                        string MadoguchiShinchokuJoukyou = data[0, 12];

                        // 中止かどうか
                        if (data[0, 11] == "3")
                        {
                            MadoguchiShinchokuJoukyou = "80";
                        }
                        else
                        {
                            // 中止でなく、報告済み
                            if (data[0, 35] == "1")
                            {
                                MadoguchiShinchokuJoukyou = "70";
                            }
                        }

                        //窓口情報（MadoguchiJouhou）
                        cmd.CommandText = "INSERT INTO MadoguchiJouhou( " +
                            "MadoguchiID " +
                            ",MadoguchiTourokuNendo " +
                            ",MadoguchiHikiwatsahi " +
                            ",MadoguchiSaishuuKensa " +
                            ",MadoguchiShouninsha " +
                            ",MadoguchiShouninnbi " +
                            ",MadoguchiShimekiribi " +
                            ",MadoguchiTourokubi " +
                            ",MadoguchiHoukokuJisshibi " +
                            ",MadoguchiChousaShubetsu " +
                            ",MadoguchiJiishiKubun " +
                            ",MadoguchiShinchokuJoukyou " +
                            ",MadoguchiJutakuBushoCD " +//受託課所支部
                            ",MadoguchiJutakuTantoushaID " +
                            ",JutakuBushoShozokuCD " +
                            ",MadoguchiTantoushaBushoCD " +
                            ",MadoguchiTantoushaCD " +
                            ",MadoguchiBushoShozokuCD " +
                            ",MadoguchiChousaKubunJibusho " + // 調査区分　自部所
                            ",MadoguchiChousaKubunShibuShibu " + //調査区分　支→支
                            ",MadoguchiChousaKubunHonbuShibu " + // 調査区分　本→支
                            ",MadoguchiChousaKubunShibuHonbu " + //調査区分　支→本
                            ",MadoguchiKanriBangou " +
                            ",MadoguchiJutakuBangou" +
                            ",MadoguchiJutakuBangouEdaban" +
                            ",MadoguchiUketsukeBangou " +
                            ",MadoguchiUketsukeBangouEdaban " +
                            ",MadoguchiHachuuKikanmei " +
                            ",MadoguchiGyoumuMeishou " +
                            ",MadoguchiKoujiKenmei " +
                            ",MadoguchiChousaHinmoku " +
                            ",MadoguchiBikou " +
                            ",MadoguchiTankaTekiyou " +
                            ",MadoguchiNiwatashi " +
                            ",MadoguchiHoukokuzumi " +
                            ",MadoguchiKanriGijutsusha " +
                            ",MadoguchiCreateDate " +
                            ",MadoguchiCreateUser " +
                            ",MadoguchiCreateProgram " +
                            ",MadoguchiUpdateDate " +
                            ",MadoguchiUpdateUser " +
                            ",MadoguchiUpdateProgram " +
                            ",MadoguchiDeleteFlag " +
                            ",MadoguchiOldBushoflg " +
                            ",MadoguchiHonbuTanpinflg " +
                            ",MadoguchiShukeiHyoFolder " +
                            ",MadoguchiHoukokuShoFolder " +
                            ",MadoguchiShiryouHolder " +
                            ",MadoguchiGyoumuKanrishaCD " +
                            ",AnkenJouhouID " +
                            ",MadoguchiHachuukikanCD " +
                            ",MadoguchiGaroonRenkei " +
                            ",MadoguchiKanryou " +
                            ",MadoguchiMitsumoriTeishutu " +
                            ",MadoguchiTeiNyuusatsu " +
                            ",MadoguchiHoukokuMale " +
                            ",MadoguchiIraiMale " +
                            ",MadoguchiIraimotoBusho " +
                            ",MadoguchiAnkenJouhouID " +
                            ",MadoguchiSaishuuKensaCheck " +
                            ",MadoguchiSystemRenban " +
                            ")VALUES(" +
                            data[0, 1] + //窓口ID
                            ",'" + data[0, 2] + "' " +  //登録年度
                            "," + data[0, 3] + " " +    //遠隔地引渡承認
                            "," + data[0, 4] + " " +             //遠隔地最終検査
                            ",N'" + data[0, 5] + "' " +          //遠隔地承認者
                            "," + data[0, 6] + " " +   //遠隔地承認日
                            "," + data[0, 7] + " " +   //調査担当者への締切日
                            "," + data[0, 8] + " " +           //登録日
                            "," + data[0, 9] + " " +   //報告実施日
                            "," + data[0, 10] + " " +   //調査種別　
                            "," + data[0, 11] + " " +   //実施区分　
                            //"," + data[0, 12] + " " +   //Madoguchi進捗状況
                            "," + MadoguchiShinchokuJoukyou + " " +   //Madoguchi進捗状況
                            ",N'" + data[0, 13] + "' " +   //受託課所支部
                            "," + data[0, 14] + " " +            //契約担当者orNULL
                            ",N'" + data[0, 15] + "' " +           //受託部所所属長の部所CD
                            ",N'" + data[0, 16] + "' " + //窓口部所
                            "," + data[0, 17] + " " +            //窓口担当者
                            ",N'" + data[0, 18] + "' " + //窓口部所所属長の部所CD
                            "," + data[0, 19] + " " + // 調査区分　自部所
                            "," + data[0, 20] + " " + //調査区分　支→支
                            "," + data[0, 21] + " " + // 調査区分　本→支
                            "," + data[0, 22] + " " + //調査区分　支→本
                            ",N'" + data[0, 23] + "' " +          //管理番号
                            ",N'" + data[0, 24].Replace("-" + data[0, 25], "") + "' " +           //受託番号
                            ",N'" + data[0, 25] + "' " +       //受託番号枝番（？）
                            ",N'" + data[0, 26] + "' " +            //特調番号
                            ",N'" + data[0, 27] + "' " +          //特調番号枝番
                            ",N'" + data[0, 28] + "' " +          //発注者名・課名
                            ",N'" + data[0, 29] + "' " +          //業務名称
                            ",N'" + data[0, 30] + "' " +          //工事件名
                            ",N'" + data[0, 31] + "' " +          //調査品目
                            ",N'" + data[0, 32] + "' " +          //備考
                            ",N'" + data[0, 33] + "' " +          //単価適用地域
                            ",N'" + data[0, 34] + "' " +          //荷渡場所
                            "," + data[0, 35] + " " +                         //0　報告済
                            ",N'" + data[0, 36] + "' " +           //管理技術者
                            ",SYSDATETIME() " +                             // 登録日時
                            ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                            ",'" + pgmName + methodName + "' " +            // 登録プログラム
                            ",SYSDATETIME() " +                             // 更新日時
                            ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                            ",'" + pgmName + methodName + "' " +            // 更新プログラム
                            ",0 " +                                         // 削除フラグ
                            ",NULL " +                              //null
                            "," + data[0, 37] + " " +               //本部単品 
                            ",N'" + data[0, 38] + "' " +          //集計表フォルダ
                            ",N'" + data[0, 39] + "' " +          //報告書フォルダ
                            ",N'" + data[0, 40] + "' " +          //調査資料フォルダ
                            "," + data[0, 41] + " " +        //業務管理者の業務管理者CD or Null
                            "," + data[0, 42] + " " +                 //AnkenJouhou.AnkenJouhouID、未受託の場合はNULL
                            ",null " +
                            "," + data[0, 43] + "" +//MadoguchiGaroonRenkei
                            ",0" + //MadoguchiKanryou
                            ",0" + //MadoguchiMitsumoriTeishutu
                            ",0" + //MadoguchiTeiNyuusatsu
                            ",0" + //MadoguchiHoukokuMale
                            ",0" + //MadoguchiIraiMale
                            ",0" + //MadoguchiIraimotoBusho
                            "," + data[0, 44] + " " +
                            ",0" +
                            "," + data[0, 45] +
                            ")";

                        cmd.ExecuteNonQuery();


                        // 案件情報から業務区分を取得する
                        string w_KyouryokuGyoumuKubun = "NULL";
                        if (data[0, 42] != "")
                        {
                            //SQL生成
                            var comboDt1 = new DataTable();
                            cmd.CommandText = "SELECT GyoumuNarabijunCD"
                                            + " FROM Mst_GyoumuKubun"
                                            + " LEFT JOIN AnkenJouhou ON GyoumuKubun = AnkenGyoumuKubunMei"
                                            + " WHERE AnkenJouhouID = " + data[0, 42];

                            //データ取得
                            var sda1 = new SqlDataAdapter(cmd);
                            sda1.Fill(comboDt1);
                            DataRow dr1 = comboDt1.Rows[0];
                            int w_AnkenGyoumuKubun = int.Parse(dr1[0].ToString());

                            // 業務区分の設定
                            if (w_AnkenGyoumuKubun != 0)
                            {
                                // GyoumuNarabijunCD の値で業務区分の値を切替
                                switch (w_AnkenGyoumuKubun)
                                {
                                    case 1: // 1:調査部（一般）
                                    case 5: // 5:事業普及部（一般）
                                    case 6: // 6:事業普及部（物品購入）
                                    case 8: // 8:総合研究所
                                        w_KyouryokuGyoumuKubun = "1";         // 1.一般受託調査
                                        break;
                                    case 3: // 3:調査部（単契）
                                        w_KyouryokuGyoumuKubun = "2";         // 2.単価契約調査
                                        break;
                                    case 4: // 4:調査部（単品)
                                        w_KyouryokuGyoumuKubun = "3";         // 3.単品契約調査
                                        break;
                                    case 2: // 2:調査部（単契含む）
                                        w_KyouryokuGyoumuKubun = "4";         // 4.単価契約を含む一般受託
                                        break;
                                    case 7: // 7:情シス部（一般契約）
                                        w_KyouryokuGyoumuKubun = "5";         // 5.情報開発受託業務
                                        break;
                                    default:
                                        w_KyouryokuGyoumuKubun = "1";         // 1.一般受託調査
                                        break;
                                }
                            }
                        }

                        //採番No（SaibanNo）を取得
                        var comboDt = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SaibanNo+SaibanCountupNo AS SaibanNo " +
                          "FROM " + "M_Saiban " +
                          "WHERE SaibanMei = 'KyouryokuIraishoID' ";

                        //データ取得
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(comboDt);
                        DataRow dr = comboDt.Rows[0];
                        int saibanKyouryokuNo = int.Parse(dr[0].ToString());

                        //採番No（SaibanNo）を更新
                        cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                            saibanKyouryokuNo + " WHERE SaibanMei = 'KyouryokuIraishoID' ";

                        cmd.ExecuteNonQuery();

                        //部所支部名、所属長CDを取得
                        var busho_comboDt = new DataTable();
                        cmd.CommandText = "SELECT"
                                        + " Mst_Busho.ShibuMei"
                                        + ", Mst_Chousain.KojinCD"
                                        + " FROM Mst_Busho"
                                        + " LEFT JOIN Mst_Chousain ON Mst_Busho.BushoShozokuChou = Mst_Chousain.ChousainMei"
                                        + " WHERE Mst_Busho.GyoumuBushoCD = '" + data[0, 16] + "'" // 窓口部所
                                        ;

                        //データ取得
                        var busho_sda = new SqlDataAdapter(cmd);
                        busho_sda.Fill(busho_comboDt);
                        DataRow busho_dr = busho_comboDt.Rows[0];
                        string bushoShibuMei = "";
                        string BushoShozokuChouCD = "";
                        if (busho_dr != null)
                        {
                            bushoShibuMei = busho_dr[0].ToString();
                            BushoShozokuChouCD = busho_dr[1].ToString();
                        }

                        //協力依頼書情報（KyouryokuIraisho）テーブル登録
                        cmd.CommandText = "INSERT INTO KyouryokuIraisho( " +
                            "KyouryokuIraishoID " +
                            ",MadoguchiID " +
                            ",KyouryokuChousaKijun " +
                            ",KyouryokuChousakijunbi " +
                            ",KyouryokuHoukokuSeigenDate " +
                            ",KyouryokuGyoumuKubun " +
                            ",KyouryokuIraiKubun " +
                            ",KyouryokuUtiawaseyouhi " +
                            ",KyouryokusakiHikiwatashi " +
                            ",KyouryokuJisshikeikakusho " +
                            //",KyourokuIraisakiTantoshaCD " +
                            ",KyouryokuGyoumuNaiyou " +
                            ",KyouryokuCreateDate " +
                            ",KyouryokuCreateUser " +
                            ",KyouryokuCreateProgram " +
                            ",KyouryokuUpdateDate" +
                            ",KyouryokuUpdateUser " +
                            ",KyouryokuUpdateProgram " +
                            ",KyouryokuDeleteFlag ";

                        // 支⇒本の場合、協力先部所は窓口部所を初期設定
                        //if (data[0, 22] == "1")
                        //{
                            cmd.CommandText += ", KyourokuIraisakiBushoOld"
                                            + ", KyourokuIraisakiTantoshaCD"
                                            ;
                        //}

                        cmd.CommandText += ")VALUES(" +
                        saibanKyouryokuNo +                    //採番
                        ",N'" + data[0, 1] + "' " +               //窓口情報
                        ",'1' " +                              //1
                        ",'" + CommonValue + "' " +            //M_COMMON_MASTER CHOUSAKIJUNBI_DEFAULT
                        "," + data[0, 7] + " " +  //調査担当者への締切日
                                                  //",NULL " +                             //空
                        "," + w_KyouryokuGyoumuKubun +          // 業務区分
                        ",NULL " +                             //空
                        ",'2' " +                              //2
                        ",'2' " +                              //2
                        ",'1' " +                              //1
                        //",NULL " +                             //NULL
                        ",'別紙の通り' " +
                        ",SYSDATETIME() " +                             // 登録日時
                        ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                        ",'" + pgmName + methodName + "' " +            // 登録プログラム
                        ",SYSDATETIME() " +                             // 更新日時
                        ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                        ",'" + pgmName + methodName + "' " +            // 更新プログラム
                        ",0 ";                                         // 削除フラグ
                        // 支⇒本の場合、協力先部所は窓口部所を初期設定
                        //if (data[0, 22] == "1")
                        //{
                            cmd.CommandText += ", N'" + bushoShibuMei + "'"
                                            + ", '" + BushoShozokuChouCD + "'"
                                            ;
                        //}
                        cmd.CommandText += ")";

                        cmd.ExecuteNonQuery();


                        //採番No（SaibanNo）を取得
                        var dt2 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SaibanNo+SaibanCountupNo AS SaibanNo " +
                          "FROM " + "M_Saiban " +
                          "WHERE SaibanMei = 'OuenUketsukeID' ";

                        //データ取得
                        var sda2 = new SqlDataAdapter(cmd);
                        sda2.Fill(dt2);

                        DataRow dr2 = dt2.Rows[0];
                        int saibanOuenNo = int.Parse(dr2[0].ToString());

                        //採番No（SaibanNo）を更新
                        cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                            saibanOuenNo + " WHERE SaibanMei = 'OuenUketsukeID' ";

                        cmd.ExecuteNonQuery();

                        //応援受付（OuenUketsuke）登録
                        cmd.CommandText = "INSERT INTO OuenUketsuke(" +
                            "OuenUketsukeID " +
                            ",MadoguchiID " +
                            ",OuenKanriNo " +
                            ",OuenCreateDate " +
                            ",OuenCreateUser " +
                            ",OuenCreateProgram " +
                            ",OuenUpdateDate " +
                            ",OuenUpdateUser " +
                            ",OuenUpdateProgram " +
                            ",OuenDeleteFlag " +
                            ")VALUES(" +
                            saibanOuenNo +
                            ", '" + data[0, 1] + "' " +
                            ",N'" + data[0, 23] + "'" +
                            ",SYSDATETIME() " +                             // 登録日時
                            ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                            ",'" + pgmName + methodName + "' " +            // 登録プログラム
                            ",SYSDATETIME() " +                             // 更新日時
                            ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                            ",'" + pgmName + methodName + "' " +            // 更新プログラム
                            ",0 " +
                            ")";
                        cmd.ExecuteNonQuery();

                        //採番No（SaibanNo）を取得
                        var tanpinDt = new DataTable();
                        int TanpinNyuuryokuID = 0;
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SaibanNo+SaibanCountupNo AS SaibanNo " +
                          "FROM " + "M_Saiban " +
                          "WHERE SaibanMei = 'TanpinNyuuryokuID' ";

                        //データ取得
                        var tanpin_sda = new SqlDataAdapter(cmd);
                        tanpin_sda.Fill(tanpinDt);

                        DataRow tanpin_dr = tanpinDt.Rows[0];
                        TanpinNyuuryokuID = int.Parse(tanpin_dr[0].ToString());

                        //採番No（TanpinNyuuryokuID）を更新
                        cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                            TanpinNyuuryokuID + " WHERE SaibanMei = 'TanpinNyuuryokuID' ";

                        cmd.ExecuteNonQuery();

                        // 単価契約の取得
                        var TankaKeiyakuDt = new DataTable();
                        int TankaKeiyakuID = 0;
                        //SQL生成
                        //cmd.CommandText = "SELECT TankaKeiyakuID FROM TankaKeiyaku WHERE AnkenJouhouID = " + ankenJouhouId;
                        cmd.CommandText = "SELECT TankaKeiyakuID FROM TankaKeiyaku"
                                        + " LEFT JOIN AnkenJouhou ON TankaKeiyaku.TankakeiyakuJutakuBangou = AnkenJouhou.AnkenJutakuBangou"
                                        + " WHERE AnkenJouhou.AnkenJouhouID = " + data[0, 42]
                                        + " ORDER BY TankaKeiyakuID DESC ";

                        //データ取得
                        var TankaKeiyaku_sda = new SqlDataAdapter(cmd);
                        TankaKeiyaku_sda.Fill(TankaKeiyakuDt);
                        if (TankaKeiyakuDt.Rows.Count > 0)
                        {
                            TankaKeiyakuID = int.Parse(TankaKeiyakuDt.Rows[0][0].ToString());
                        }

                        ////単品入力（TanpinNyuuryoku）登録
                        //cmd.CommandText = "INSERT INTO TanpinNyuuryoku(" +
                        //  "TanpinNyuuryokuID " +
                        //  ",MadoguchiID " +
                        //  //",TanpinGyoumuCD " +
                        //  ",TanpinCreateDate " +
                        //  ",TanpinCreateUser " +
                        //  ",TanpinCreateProgram " +
                        //  ",TanpinUpdateDate " +
                        //  ",TanpinUpdateUser " +
                        //  ",TanpinUpdateProgram " +
                        //  ",TanpinDeleteFlag " +
                        //  ")VALUES(" +
                        //  TanpinNyuuryokuID +
                        //  ", '" + data[0, 1] + "' " +
                        //  //", " + "0" + " " + //確定で0
                        //  ",SYSDATETIME() " +                             // 登録日時
                        //  ",'" + UserInfos[0] + "' " +                    // 登録ユーザ
                        //  ",'" + pgmName + methodName + "' " +            // 登録プログラム
                        //  ",SYSDATETIME() " +                             // 更新日時
                        //  ",'" + UserInfos[0] + "' " +                    // 更新ユーザ
                        //  ",'" + pgmName + methodName + "' " +            // 更新プログラム
                        //  ",0 " +                                         // 削除フラグ
                        //  ")";
                        //cmd.ExecuteNonQuery();

                        // No352:単品入力画面に、部署、役職、担当者、電話、FAX、メールがコピーされない。対応
                        cmd.CommandText = "INSERT INTO TanpinNyuuryoku(" +
                          "TanpinNyuuryokuID " +
                          ",MadoguchiID " +
                          ",TanpinGyoumuCD " +
                          ",TanpinHachuubusho " +
                          ",TanpinYakushoku " +
                          ",TanpinHachuuTantousha " +
                          ",TanpinTel " +
                          ",TanpinFax " +
                          ",TanpinMail " +
                          ",TanpinCreateDate " +
                          ",TanpinCreateUser " +
                          ",TanpinCreateProgram " +
                          ",TanpinUpdateDate " +
                          ",TanpinUpdateUser " +
                          ",TanpinUpdateProgram " +
                          ",TanpinDeleteFlag " +
                          ")VALUES(" +
                          TanpinNyuuryokuID +
                          ", '" + data[0, 1] + "' " +
                          ", " + TankaKeiyakuID + " " +
                          // 新規には調査概要 別タブは存在しないので、コメントアウト
                          //", '" + item6_TanpinHachuubusho.Text + "' " +
                          //", '" + item6_TanpinYakushoku.Text + "' " +
                          //", '" + item6_TanpinHachuuTantousha.Text + "' " +
                          //", '" + item6_TanpinTel.Text + "' " +
                          //", '" + item6_TanpinFax.Text + "' " +
                          //", '" + item6_TanpinMail.Text + "' " +
                          ", '' " +
                          ", '' " +
                          ", '' " +
                          ", '' " +
                          ", '' " +
                          ", '' " +
                          ",SYSDATETIME() " +                             // 登録日時
                          ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                          ",'" + pgmName + methodName + "' " +            // 登録プログラム
                          ",SYSDATETIME() " +                             // 更新日時
                          ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                          ",'" + pgmName + methodName + "' " +            // 更新プログラム
                          ",0 " +                                         // 削除フラグ
                          ")";
                        cmd.ExecuteNonQuery();


                        ////採番No（SaibanNo）を取得
                        //var dt3 = new DataTable();
                        ////SQL生成
                        //cmd.CommandText = "SELECT " +
                        //  "SaibanNo+SaibanCountupNo AS SaibanNo " +
                        //  "FROM " + "M_Saiban " +
                        //  "WHERE SaibanMei = 'HistoryID' ";

                        ////データ取得
                        //var sda3 = new SqlDataAdapter(cmd);
                        //sda3.Fill(dt3);

                        //DataRow dr3 = dt3.Rows[0];
                        //int saibanHistoryNo = int.Parse(dr3[0].ToString());

                        ////採番No（SaibanNo）を更新
                        //cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                        //    saibanHistoryNo + " WHERE SaibanMei = 'HistoryID' ";

                        //cmd.ExecuteNonQuery();
                        
                        ////履歴登録
                        //cmd.CommandText = "INSERT INTO T_HISTORY(" +
                        //    "H_DATE_KEY " +
                        //    ",H_NO_KEY " +
                        //    ",H_OPERATE_DT " +
                        //    ",H_OPERATE_USER_ID " +
                        //    ",H_OPERATE_USER_MEI " +
                        //    ",H_OPERATE_USER_BUSHO_CD " +
                        //    ",H_OPERATE_USER_BUSHO_MEI " +
                        //    ",H_OPERATE_NAIYO " +
                        //    ",H_ProgramName " +
                        //    ",MadoguchiID " +
                        //    ",HistoryBeforeTantoubushoCD " +
                        //    ",HistoryBeforeTantoushaCD " +
                        //    ",HistoryAfterTantoubushoCD " +
                        //    ",HistoryAfterTantoushaCD " +
                        //    ")VALUES(" +
                        //    "SYSDATETIME() " + 
                        //    ", " + saibanHistoryNo + " " +
                        //    ",SYSDATETIME() " +
                        //    ",'" + UserInfos[0] + "' " +
                        //    ",'" + UserInfos[1] + "' " +
                        //    ",'" + UserInfos[2] + "' " +
                        //    ",'" + UserInfos[3] + "' " +
                        //    ",'調査概要を追加しました ID:" + data[0, 1] + " Garoon連携区分:" + data[0, 43] + "' " +
                        //    ",'" + pgmName + methodName + "' " +
                        //    "," + data[0, 1] + " " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ")";

                        Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "調査概要を追加しました ID:" + data[0, 1] + " Garoon連携区分:" + data[0, 43], pgmName + methodName, data[0, 1]);


                        //cmd.ExecuteNonQuery();

                        //既存のGaroon宛先追加処理
                        
                        //窓口情報（MadoguchiJouhou）テーブルからGaroon連携対象（GaroonRenkeiKubn）を取得
                        var dt4 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "MadoguchiTantoushaCD,MadoguchiKanriGijutsusha " +
                          ",MadoguchiUketsukeBangou,MadoguchiUketsukeBangouEdaban,MadoguchiGaroonRenkei " +
                          "FROM MadoguchiJouhou " +
                          "WHERE MadoguchiID = " + data[0, 1] + "";

                        //データ取得
                        var sda4 = new SqlDataAdapter(cmd);
                        sda4.Fill(dt4);

                        String atesaki = dt4.Rows[0][0].ToString();
                        String kanriGijutusha = dt4.Rows[0][1].ToString();
                        String tokuchouNo = dt4.Rows[0][2].ToString();
                        String tokuchouNoEda = dt4.Rows[0][3].ToString();
                        String garoonOn = dt4.Rows[0][4].ToString();

                        //窓口メール送信（MadoguchiMail）テーブルからメッセージID（MadoguchiMailMessageID）を取得
                        var dt5 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "MadoguchiMailMessageID " +
                          "FROM MadoguchiMail " +
                          //"WHERE MadoguchiMailTokuchoBangou = '" + tokuchouNo + "-" + tokuchouNoEda + "'" +
                          "WHERE MadoguchiMailTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + tokuchouNo + "'" +
                          "AND MadoguchiMailTokuchoBangouEda COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + data[0, 27] + "' ";

                        //データ取得
                        var sda5 = new SqlDataAdapter(cmd);
                        sda5.Fill(dt5);

                        // 863 NULLではなく、0が正しい
                        String mailMessageID = "0";

                        if (dt5.Rows.Count > 0)
                        {
                            mailMessageID = dt5.Rows[0][0].ToString();
                        }

                        //管理技術者（MadoguchiJouhou.MadoguchiKanriGijutsusha）が空でない場合
                        if (!String.IsNullOrEmpty(kanriGijutusha) && kanriGijutusha != "0")
                        {
                            //宛先がnullじゃない
                            if (!String.IsNullOrEmpty(atesaki))
                            {
                                atesaki = atesaki + ";" + kanriGijutusha;
                            }
                            //宛先がnull
                            else
                            {
                                atesaki += kanriGijutusha;
                            }
                        }

                        //MadoguchiJouhouMadoguchiL1Chouを取得
                        var dt6 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT DISTINCT " +
                          "MadoguchiL1ChousaTantoushaCD,MadoguchiL1ChousaBushoCD " +
                          "FROM MadoguchiJouhouMadoguchiL1Chou " +
                          "WHERE MadoguchiID=" + data[0, 1] + "";

                        //データ取得
                        var sda6 = new SqlDataAdapter(cmd);
                        sda6.Fill(dt6);

                        for (int i = 0; i < dt6.Rows.Count; i++)
                        {
                            //調査員担当者（MadoguchiJouhouMadoguchiL1Chou.MadoguchiL1ChousaTantoushaCD）が空でない場合
                            String chousaTantousha = dt6.Rows[i][0].ToString();
                            if (!String.IsNullOrEmpty(chousaTantousha) && chousaTantousha != "0")
                            {
                                //宛先が空でない場合
                                if (!String.IsNullOrEmpty(atesaki))
                                {
                                    atesaki = atesaki + ";" + chousaTantousha;
                                }
                                //宛先がnull
                                else
                                {
                                    atesaki = chousaTantousha;
                                }
                            }

                            //調査担当部所コード（MadoguchiJouhouMadoguchiL1Chou.MadoguchiL1ChousaBushoCD）が空でない場合
                            String chousaTantoubusho = dt6.Rows[i][1].ToString();
                            if (!String.IsNullOrEmpty(chousaTantoubusho))
                            {
                                //支部応援（Mst_Shibuouen）と、調査員マスタ（Mst_Chousain）を結合し担当者を取得
                                var dt7 = new DataTable();
                                //SQL生成
                                cmd.CommandText = "SELECT " +
                                  "Mst_Chousain.KojinCD " +
                                  "FROM Mst_Shibuouen INNER JOIN Mst_Chousain ON " +
                                  "Mst_Chousain.KojinCD = Mst_Shibuouen.ShibuouenKojinCD " +
                                  //"AND Mst_Shibuouen.ShibuouenDeleteFlag = 0 " +
                                  //"AND Mst_Chousain.RetireFLG = 0 " +
                                  "AND Mst_Chousain.GyoumuBushoCD ='" + chousaTantoubusho + "' ";

                                //データ取得
                                var sda7 = new SqlDataAdapter(cmd);
                                sda7.Fill(dt7);

                                for (int j = 0; j < dt7.Rows.Count; j++)
                                {
                                    if (dt7.Rows[j][0] != null && dt7.Rows[j][0].ToString() != "0")
                                    {
                                        //宛先が空でない場合
                                        if (!String.IsNullOrEmpty(atesaki))
                                        {
                                            atesaki = atesaki + ";" + dt7.Rows[j][0].ToString();
                                        }
                                        //宛先がnull
                                        else
                                        {
                                            atesaki = dt7.Rows[j][0].ToString();
                                        }
                                    }
                                }//for end
                            }//if end
                        }//for end

                        //宛先が空でない場合
                        if (!String.IsNullOrEmpty(atesaki))
                        {
                            //メール情報CSV抽出用ワーク（MailInfoCSVWork）テーブルから
                            var dt8 = new DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "MailInfoCSVWorkID " +
                              "FROM MailInfoCSVWork " +
                              "WHERE MailInfoCSVWorkTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + tokuchouNo + "-" + tokuchouNoEda + "' " +
                              "AND MailInfoCSVWorkCSVOutFlg = 0 " +
                              "AND MailInfoCSVWorkGaRenkeiFlg = 0 " +
                              "AND MailInfoCSVWorkDeleteFlag = 0";

                            //データ取得
                            var sda8 = new SqlDataAdapter(cmd);
                            sda8.Fill(dt8);

                            //メール情報CSV抽出用ワークのデータがある
                            if (dt8.Rows.Count > 0)
                            {
                                String workId = dt8.Rows[0][0].ToString();
                                //ガルーン連携がチェックの場合
                                if ("1".Equals(garoonOn))
                                {

                                    //メール情報CSV抽出用ワーク（MailInfoCSVWork）テーブル更新
                                    cmd.CommandText = "UPDATE MailInfoCSVWork SET " +
                                        "MailInfoCSVWorkAtesakiUser = '" + atesaki + "' " +
                                        ",MailInfoCSVWorkUpdateDate = SYSDATETIME() " +
                                        ",MailInfoCSVWorkUpdateUser = N'" + UserInfos[0] + "' " +
                                        ",MailInfoCSVWorkUpdateProgram = '窓口ミハル' " +
                                        "WHERE MailInfoCSVWorkTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + tokuchouNo + "-" + tokuchouNoEda + "' AND MailInfoCSVWorkDeleteFlag = 0";

                                    cmd.ExecuteNonQuery();

                                }
                                //連携が未チェックの場合
                                else
                                {

                                    //メール情報CSV抽出用ワーク（MailInfoCSVWork）テーブル削除
                                    cmd.CommandText = "DELETE FROM MailInfoCSVWork " +
                                        "WHERE MailInfoCSVWorkTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + tokuchouNo + "-" + tokuchouNoEda + "' ";

                                    cmd.ExecuteNonQuery();
                                }
                            }
                            //メール情報CSV抽出用ワークのデータがない
                            else
                            {
                                //ガルーン連携がチェックの場合
                                if ("1".Equals(garoonOn))
                                {
                                    //採番No（SaibanNo）を取得
                                    var dt9 = new DataTable();
                                    //SQL生成
                                    cmd.CommandText = "SELECT " +
                                      "SaibanNo+SaibanCountupNo AS SaibanNo " +
                                      "FROM " + "M_Saiban " +
                                      "WHERE SaibanMei = 'MailInfoCSVWorkID' ";

                                    //データ取得
                                    var sda9 = new SqlDataAdapter(cmd);
                                    sda9.Fill(dt9);

                                    int saibanMailInfoNo = int.Parse(dt9.Rows[0][0].ToString());

                                    //採番No（SaibanNo）を更新
                                    cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                                        saibanMailInfoNo + " WHERE SaibanMei = 'MailInfoCSVWorkID' ";

                                    cmd.ExecuteNonQuery();

                                    //メール情報CSV抽出用ワーク（MailInfoCSVWork）テーブル登録
                                    cmd.CommandText = "INSERT INTO MailInfoCSVWork(" +
                                    "MailInfoCSVWorkID " +
                                    ",MailInfoCSVWorkMadoguchiID " +
                                    ",MailInfoCSVWorkTokuchoBangou " +
                                    ",MailInfoCSVWorkMessageID " +
                                    ",MailInfoCSVWorkAtesakiUser " +
                                    ",MailInfoCSVWorkCSVOutFlg " +
                                    ",MailInfoCSVWorkGaRenkeiFlg " +
                                    ",MailInfoCSVWorkCreateDate " +
                                    ",MailInfoCSVWorkCreateUser " +
                                    ",MailInfoCSVWorkCreateProgram " +
                                    ",MailInfoCSVWorkDeleteFlag" +
                                    ")VALUES(" +
                                    saibanMailInfoNo +
                                    ", '" + data[0, 1] + "' " +
                                    ",N'" + tokuchouNo + "-" + tokuchouNoEda + "' " +
                                    "," + mailMessageID + " " +
                                    ",'" + atesaki + "' " +
                                    ",0" +
                                    ",0" +
                                    ",SYSDATETIME()" +
                                    ",N'" + UserInfos[0] + "'" +
                                    ",'窓口ミハル'" +
                                    ",0" +
                                    ")";
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                        //宛先がnull
                        else
                        {

                            //メール情報CSV抽出用ワーク（MailInfoCSVWork）テーブル削除
                            cmd.CommandText = "DELETE FROM MailInfoCSVWork " +
                                "WHERE MailInfoCSVWorkTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + tokuchouNo + "-" + tokuchouNoEda + "' " +
                                "AND MailInfoCSVWorkCSVOutFlg = 0 " +
                                "AND MailInfoCSVWorkGaRenkeiFlg = 0 ";

                            cmd.ExecuteNonQuery();

                        }
                        


                        //// GaroonTsuikaAtesakiに窓口部所に一致する支部応援のユーザーを追加する
                        //var garoonDt = new DataTable();
                        ////SQL生成
                        //cmd.CommandText = "SELECT " +
                        //  "Mst_Chousain.KojinCD " +
                        //  ",Mst_Chousain.ChousainMei " +
                        //  ",Mst_Chousain.GyoumuBushoCD " +
                        //  ",mb.BushokanriboKamei " +
                        //  "FROM Mst_Shibuouen INNER JOIN Mst_Chousain ON " +
                        //  "Mst_Chousain.KojinCD = Mst_Shibuouen.ShibuouenKojinCD " +
                        //  "AND Mst_Shibuouen.ShibuouenDeleteFlag = 0 " +
                        //  "AND Mst_Chousain.RetireFLG = 0 " +
                        //  "AND Mst_Chousain.GyoumuBushoCD ='" +data[0,16] + "' " +
                        //  "INNER JOIN Mst_Busho mb ON mb.GyoumuBushoCD = Mst_Chousain.GyoumuBushoCD ";
                        ////データ取得
                        //var garronSda = new SqlDataAdapter(cmd);
                        //garronSda.Fill(garoonDt);

                        string KojinCD = "";
                        string ChousainMei = "";
                        string GyoumuBushoCD = "";
                        string BushoMei = "";

                        //if (garoonDt != null && garoonDt.Rows.Count > 0)
                        //{
                        //    for (int i = 0; i < garoonDt.Rows.Count; i++)
                        //    {
                        //        KojinCD = garoonDt.Rows[i][0].ToString();
                        //        ChousainMei = garoonDt.Rows[i][1].ToString();
                        //        GyoumuBushoCD = garoonDt.Rows[i][2].ToString();
                        //        BushoMei = garoonDt.Rows[i][3].ToString();

                        //        // GaroonTsuikaAtesakiに登録
                        //        cmd.CommandText = "INSERT INTO GaroonTsuikaAtesaki ( " +
                        //        " GaroonTsuikaAtesakiID " +
                        //        ",GaroonTsuikaAtesakiMadoguchiID " +
                        //        ",GaroonTsuikaAtesakiBushoCD " +
                        //        ",GaroonTsuikaAtesakiBusho " +
                        //        ",GaroonTsuikaAtesakiTantoushaCD " +
                        //        ",GaroonTsuikaAtesakiTantousha " +
                        //        ",GaroonTsuikaAtesakiCreateDate " +
                        //        ",GaroonTsuikaAtesakiCreateUser " +
                        //        ",GaroonTsuikaAtesakiCreateProgram " +
                        //        ",GaroonTsuikaAtesakiUpdateDate " +
                        //        ",GaroonTsuikaAtesakiUpdateUser " +
                        //        ",GaroonTsuikaAtesakiUpdateProgram " +
                        //        ",GaroonTsuikaAtesakiDeleteFlag " +
                        //        ") VALUES (" +
                        //        "'" + getSaiban("GaroonTsuikaAtesakiID") + "' " + // GaroonTsuikaAtesakiID
                        //        ",'" + data[0,1] + "' " +          // GaroonTsuikaAtesakiMadoguchiID
                        //        ",'" + GyoumuBushoCD + "' " +              // GaroonTsuikaAtesakiBushoCD
                        //        ",'" + BushoMei + "' " +                   // GaroonTsuikaAtesakiBusho
                        //        ",'" + KojinCD + "' " +                    // GaroonTsuikaAtesakiTantoushaCD
                        //        ",'" + ChousainMei + "' " +                // GaroonTsuikaAtesakiTantousha
                        //        ",SYSDATETIME() " +                        // GaroonTsuikaAtesakiCreateDate
                        //        ",'" + UserInfos[0] + "'" +                // GaroonTsuikaAtesakiCreateUser
                        //        ",'窓口ミハル'" +                          // GaroonTsuikaAtesakiCreateProgram
                        //        ",SYSDATETIME() " +                        // GaroonTsuikaAtesakiUpdateDate
                        //        ",'" + UserInfos[0] + "'" +                // GaroonTsuikaAtesakiUpdateUser
                        //        ",'窓口ミハル'" +                          // GaroonTsuikaAtesakiUpdateProgram
                        //        ",0 " +                                    // GaroonTsuikaAtesakiDeleteFlag
                        //        ") ";

                        //        cmd.ExecuteNonQuery();
                        //    }
                        //}

                        // 469 Garoon連携担当者の自動設定
                        // 担当者宛先の自動設定（窓口担当者、調査担当者除く）

                        // 管理技術者が空でない場合、GaroonTsuikaAtesakiテーブルにデータを追加する
                        if (data[0, 36] != null && data[0, 36] != "")
                        {
                            var dt9 = new DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "KojinCD " +
                              ",ChousainMei " +
                              ",mc.GyoumuBushoCD " +
                              ",mb.ShibuMei + ' ' + IsNull(mb.KaMei,'') AS BushoMei " +
                              "FROM Mst_Chousain mc " +
                              "LEFT JOIN Mst_Busho mb ON mb.GyoumuBushoCD = mc.GyoumuBushoCD " +
                              "WHERE mc.KojinCD = '" + data[0, 36] + "' ";

                            //データ取得
                            var sda9 = new SqlDataAdapter(cmd);
                            sda9.Fill(dt9);

                            KojinCD = "";
                            ChousainMei = "";
                            GyoumuBushoCD = "";
                            BushoMei = "";

                            if (dt9 != null && dt9.Rows.Count > 0)
                            {
                                KojinCD = dt9.Rows[0][0].ToString();
                                ChousainMei = dt9.Rows[0][1].ToString();
                                GyoumuBushoCD = dt9.Rows[0][2].ToString();
                                BushoMei = dt9.Rows[0][3].ToString();

                                // 窓口部所に一致する支部応援で既に登録済みかもしれないので、存在チェック
                                string where = "GaroonTsuikaAtesakiMadoguchiID = '" + data[0, 1] + "' " +
                                               "AND GaroonTsuikaAtesakiBushoCD = '" + GyoumuBushoCD + "' " +
                                               "AND GaroonTsuikaAtesakiTantoushaCD = '" + KojinCD + "' " +
                                               "AND GaroonTsuikaAtesakiDeleteFlag <> 1";

                                var tmpdt = getData("KojinCD", "KojinCD", "GaroonTsuikaAtesaki", where);
                                // データ件数が0件なら登録
                                if (tmpdt != null && tmpdt.Rows.Count == 0)
                                {
                                    // GaroonTsuikaAtesakiに登録
                                    cmd.CommandText = "INSERT INTO GaroonTsuikaAtesaki ( " +
                                    " GaroonTsuikaAtesakiID " +
                                    ",GaroonTsuikaAtesakiMadoguchiID " +
                                    ",GaroonTsuikaAtesakiBushoCD " +
                                    ",GaroonTsuikaAtesakiBusho " +
                                    ",GaroonTsuikaAtesakiTantoushaCD " +
                                    ",GaroonTsuikaAtesakiTantousha " +
                                    ",GaroonTsuikaAtesakiCreateDate " +
                                    ",GaroonTsuikaAtesakiCreateUser " +
                                    ",GaroonTsuikaAtesakiCreateProgram " +
                                    ",GaroonTsuikaAtesakiUpdateDate " +
                                    ",GaroonTsuikaAtesakiUpdateUser " +
                                    ",GaroonTsuikaAtesakiUpdateProgram " +
                                    ",GaroonTsuikaAtesakiDeleteFlag " +
                                    ") VALUES (" +
                                    "'" + getSaiban("GaroonTsuikaAtesakiID") + "' " + // GaroonTsuikaAtesakiID
                                    ",'" + data[0, 1] + "' " +                 // GaroonTsuikaAtesakiMadoguchiID
                                    ",'" + GyoumuBushoCD + "' " +              // GaroonTsuikaAtesakiBushoCD
                                    ",N'" + BushoMei + "' " +                   // GaroonTsuikaAtesakiBusho
                                    ",'" + KojinCD + "' " +                    // GaroonTsuikaAtesakiTantoushaCD
                                    ",N'" + ChousainMei + "' " +                // GaroonTsuikaAtesakiTantousha
                                    ",SYSDATETIME() " +                             // 登録日時
                                    ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                                    ",'" + pgmName + methodName + "' " +            // 登録プログラム
                                    ",SYSDATETIME() " +                             // 更新日時
                                    ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                                    ",'" + pgmName + methodName + "' " +            // 更新プログラム
                                    ",0 " +                                         // 削除フラグ
                                    ") ";

                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }

                        mes += GetMessage("I20101", "");

                    }
                    //更新の場合
                    else
                    {
                        string MadoguchiTourokubi = "";
                        string MadoguchiJutakuBangou = "";
                        string MadoguchiJutakuBangouEdaban = "";
                        string MadoguchiUketsukeBangou = "";
                        string MadoguchiUketsukeBangouEdaban = "";
                        string AnkenJouhouID = "";

                        // 変更前の登録日、受託番号、受託番号枝番を取っておく
                        var dtMadoguchi = new DataTable();
                        cmd.CommandText = "SELECT MadoguchiTourokubi,MadoguchiJutakuBangou,MadoguchiJutakuBangouEdaban,MadoguchiUketsukeBangou,MadoguchiUketsukeBangouEdaban,AnkenJouhouID " +
                            "FROM MadoguchiJouhou " +
                            "WHERE MadoguchiID = " + MadoguchiID + "";
                        //データ取得
                        var madoguchiSda = new SqlDataAdapter(cmd);
                        madoguchiSda.Fill(dtMadoguchi);

                        if(dtMadoguchi != null && dtMadoguchi.Rows.Count > 0)
                        {
                            MadoguchiTourokubi = dtMadoguchi.Rows[0][0].ToString();
                            MadoguchiJutakuBangou = dtMadoguchi.Rows[0][1].ToString();
                            MadoguchiJutakuBangouEdaban = dtMadoguchi.Rows[0][2].ToString();
                            MadoguchiUketsukeBangou = dtMadoguchi.Rows[0][3].ToString();
                            MadoguchiUketsukeBangouEdaban = dtMadoguchi.Rows[0][4].ToString();
                            AnkenJouhouID = dtMadoguchi.Rows[0][5].ToString();
                        }

                        // 中止フラグ 1:中止
                        int madoguchiChuushi = 0;

                        //進捗状況 10：依頼　調査中→40：集計中　70：二次検済
                        int shinchoku = 10;
                        //実施区分が1:実施
                        if ("1".Equals(data[0, 11]))
                        {
                            //報告済み
                            if ("1".Equals(data[0, 34]))
                            {
                                // 70：二次検済
                                shinchoku = 70;
                                // 中止フラグを0
                                madoguchiChuushi = 0;
                            }
                            //報告済みじゃない
                            else
                            {
                                //最小値を取る
                                var dtShinchoku = new DataTable();
                                cmd.CommandText = "SELECT min(MadoguchiL1ChousaShinchoku)as minShonchoku " +
                                    "FROM MadoguchiJouhouMadoguchiL1Chou " +
                                    "WHERE MadoguchiID = " + MadoguchiID + "";
                                //データ取得
                                var sda = new SqlDataAdapter(cmd);
                                sda.Fill(dtShinchoku);

                                String minStr = dtShinchoku.Rows[0][0].ToString();
                                if (minStr != "NULL" && minStr != "null" && !String.IsNullOrEmpty(minStr))
                                {
                                    int minShinchoku = int.Parse(minStr);

                                    //if (minShinchoku == 10)
                                    //{
                                    //    shinchoku = 40;
                                    //}
                                    //else
                                    //{
                                    //    shinchoku = 10;
                                    //}
                                    shinchoku = minShinchoku;
                                    //// 進捗が中止が最小だった場合、2次検済みとする
                                    //if (shinchoku == 80)
                                    //{
                                    //    shinchoku = 70;
                                    //}
                                }
                                else
                                {
                                    shinchoku = 10;
                                }
                            }
                        }
                        //実施区分が2打診中
                        else if ("2".Equals(data[0, 11]))
                        {
                            // 中止フラグを0
                            madoguchiChuushi = 0;
                            //報告済み
                            if ("1".Equals(data[0, 34]))
                            {
                                shinchoku = 70;
                            }
                            //報告済みじゃない
                            else
                            {
                                //最小値を取る
                                var dtShinchoku = new DataTable();
                                cmd.CommandText = "SELECT min(MadoguchiL1ChousaShinchoku) " +
                                    "FROM MadoguchiJouhouMadoguchiL1Chou " +
                                    "WHERE MadoguchiID = " + MadoguchiID + "";
                                //データ取得
                                var sda = new SqlDataAdapter(cmd);
                                sda.Fill(dtShinchoku);

                                string minStr = dtShinchoku.Rows[0][0].ToString();
                                if (minStr != "NULL" && minStr != "null" && !String.IsNullOrEmpty(minStr))
                                {
                                    shinchoku = int.Parse(minStr);
                                    // 進捗が中止が最小だった場合、2次検済みとする
                                    //if(shinchoku == 80)
                                    //{
                                    //    shinchoku = 70;
                                    //}
                                }
                                else
                                {
                                    shinchoku = 10;
                                }
                            }
                        }
                        // 実施区分が3:中止
                        else if ("3".Equals(data[0, 11]))
                        {
                            // 中止フラグを1
                            madoguchiChuushi = 1;
                            // 進捗を中止
                            shinchoku = 80;
                        }

                        // 担当部所で持つ親情報（登録年度、特調番号（受付番号・枝番）、発注者名・課名、集計表リンク）に変更があれば、
                        // 担当部所の親データを更新する
                        string tourokuNendoOld = "";
                        string UketsukeBangouOld = "";
                        string UketsukeBangouEdabanOld = "";
                        string HachuuKikanmeiOld = "";
                        string ShukeiHyoFolderOld = "";

                        // GaroonTsuikaAtesakiに追加するか判断の為に使用
                        string MadoguchiTantoushaBushoCD = "";
                        // 親データ更新フラグ true:更新する false:更新しない
                        Boolean parentUpdateFlg = false;

                        var dtParent = new DataTable();
                        cmd.CommandText = "SELECT " +
                            " MadoguchiTourokuNendo " +
                            ",MadoguchiUketsukeBangou " +
                            ",MadoguchiUketsukeBangouEdaban " +
                            ",MadoguchiHachuuKikanmei " +
                            ",MadoguchiShukeiHyoFolder " +
                            ",MadoguchiTantoushaBushoCD " +
                            "FROM MadoguchiJouhou " +
                            "WHERE MadoguchiID = " + MadoguchiID + "";
                        //データ取得
                        var sdaParent = new SqlDataAdapter(cmd);
                        sdaParent.Fill(dtParent);

                        if (dtParent != null && dtParent.Rows.Count > 0)
                        {
                            tourokuNendoOld = dtParent.Rows[0][0].ToString();
                            UketsukeBangouOld = dtParent.Rows[0][1].ToString();
                            UketsukeBangouEdabanOld = dtParent.Rows[0][2].ToString();
                            HachuuKikanmeiOld = dtParent.Rows[0][3].ToString();
                            ShukeiHyoFolderOld = dtParent.Rows[0][4].ToString();

                            MadoguchiTantoushaBushoCD = dtParent.Rows[0][5].ToString();

                            // 登録年度・・・・・・・・0,2
                            // 特調番号（受付番号）・・0,25
                            // 特調番号（枝番）・・・・0,26
                            // 発注者名・課名・・・・・0,27
                            // 集計表リンク・・・・・・0,37
                            if (tourokuNendoOld == data[0, 2]
                                && UketsukeBangouOld == data[0, 25]
                                && UketsukeBangouEdabanOld == data[0, 26]
                                && HachuuKikanmeiOld == data[0, 27]
                                && ShukeiHyoFolderOld == data[0, 37]
                                )
                            {
                                // 全部一致なので、フラグは立てない
                            }
                            else
                            {
                                parentUpdateFlg = true;
                            }
                        }

                        //窓口情報更新
                        cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                            "MadoguchiTourokuNendo = " + data[0, 2] + " " +
                            ",MadoguchiHikiwatsahi = " + data[0, 3] + " " +
                            ",MadoguchiSaishuuKensa = " + data[0, 4] + " " +
                            ",MadoguchiShouninsha = N'" + data[0, 5] + "' " +
                            ",MadoguchiShouninnbi = " + data[0, 6] + " " +
                            ",MadoguchiShimekiribi = " + data[0, 7] + " " +
                            ",MadoguchiTourokubi = " + data[0, 8] + " " +
                            ",MadoguchiHoukokuJisshibi = " + data[0, 9] + " " +
                            ",MadoguchiChousaShubetsu = " + data[0, 10] + " " +
                            ",MadoguchiJiishiKubun = " + data[0, 11] + " " +
                            ",MadoguchiShinchokuJoukyou = " + shinchoku + " " +
                            ",MadoguchiJutakuBushoCD = '" + data[0, 12] + "' " +
                            ",MadoguchiJutakuTantoushaID = " + data[0, 13] + " " +            //契約担当者orNULL　
                            ",JutakuBushoShozokuCD = '" + data[0, 14] + "' " +
                            ",MadoguchiTantoushaBushoCD = '" + data[0, 15] + "' " +
                            ",MadoguchiTantoushaCD = " + data[0, 16] + " " +
                            ",MadoguchiBushoShozokuCD = '" + data[0, 17] + "' " +
                            ",MadoguchiChousaKubunJibusho = " + data[0, 18] + " " +
                            ",MadoguchiChousaKubunShibuShibu = " + data[0, 19] + " " +
                            ",MadoguchiChousaKubunHonbuShibu = " + data[0, 20] + " " +
                            ",MadoguchiChousaKubunShibuHonbu = " + data[0, 21] + " " +
                            ",MadoguchiKanriBangou = N'" + data[0, 22] + "' " +
                            ",MadoguchiJutakuBangou = N'" + data[0, 23].Replace("-" + data[0, 24], "") + "' " +
                            ",MadoguchiJutakuBangouEdaban = N'" + data[0, 24] + "' " +       //受託番号枝番
                            ",MadoguchiUketsukeBangou = N'" + data[0, 25] + "' " + //特調番号
                            ",MadoguchiUketsukeBangouEdaban = N'" + data[0, 26] + "' " +
                            ",MadoguchiHachuuKikanmei = N'" + data[0, 27] + "' " +
                            ",MadoguchiGyoumuMeishou = N'" + data[0, 28] + "' " +
                            ",MadoguchiKoujiKenmei = N'" + data[0, 29] + "' " +
                            ",MadoguchiChousaHinmoku = N'" + data[0, 30] + "' " +
                            ",MadoguchiBikou = N'" + data[0, 31] + "' " +
                            ",MadoguchiTankaTekiyou = N'" + data[0, 32] + "' " +
                            ",MadoguchiNiwatashi = N'" + data[0, 33] + "' " +
                            ",MadoguchiHoukokuzumi = N'" + data[0, 34] + "' " +
                            ",MadoguchiKanriGijutsusha = N'" + data[0, 35] + "' " +
                            ",MadoguchiUpdateDate = SYSDATETIME()" +
                            ",MadoguchiUpdateUser = '" + UserInfos[0] + "' " +
                            ",MadoguchiUpdateProgram = '" + pgmName + methodName + "' " +
                            ",MadoguchiDeleteFlag = 0 " +
                            ",MadoguchiOldBushoflg = 0 " +
                            ",MadoguchiHonbuTanpinflg = " + data[0, 36] + " " +
                            ",MadoguchiShukeiHyoFolder  = N'" + data[0, 37] + "' " +
                            ",MadoguchiHoukokuShoFolder = N'" + data[0, 38] + "' " +
                            ",MadoguchiShiryouHolder = N'" + data[0, 39] + "' " +
                            ",MadoguchiGyoumuKanrishaCD = " + data[0, 40] + " " + //業務管理者の業務管理者CD or Null
                            ",AnkenJouhouID = " + data[0, 41] + " " +
                            ",MadoguchiGaroonRenkei = " + data[0, 44] + " " +
                            ",MadoguchiHachuukikanCD = NULL "　+
                            ",MadoguchiChuushi = " + madoguchiChuushi; // 中止フラグ

                        //受託番号(＝案件番号,特調番号)が変わった場合
                        if (!"NULL".Equals(data[0, 42]))
                        {
                            cmd.CommandText += ",MadoguchiSystemRenban = " + data[0, 42] + " ";
                        }

                        cmd.CommandText += " WHERE MadoguchiID =  " + data[0, 1];
                        cmd.ExecuteNonQuery();


                        //部所ＣＤから部所支部と部所課名を取得
                        var bushoDt = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "ShibuMei,KaMei " +
                          "FROM Mst_Busho " +
                          "WHERE GyoumuBushoCD = '" + data[0, 15] + "' ";

                        //データ取得
                        var sdaBusho = new SqlDataAdapter(cmd);
                        sdaBusho.Fill(bushoDt);


                        //GyoumuJouhouMadoguchiの窓口情報を更新
                        cmd.CommandText = "UPDATE GyoumuJouhouMadoguchi  SET " +
                            " GyoumuJouhouMadoGyoumuBushoCD = " + "'" + data[0, 15] + "' " +
                            ", GyoumuJouhouMadoShibuMei = " + "N'" + bushoDt.Rows[0][0] + "' " +
                            ", GyoumuJouhouMadoKamei = " + "N'" + bushoDt.Rows[0][1] + "' " +
                            ", GyoumuJouhouMadoKojinCD = " + "'" + data[0, 16] + "' " +
                            ", GyoumuJouhouMadoChousainMei = " + "N'" + data[0, 43] + "' " +
                            " WHERE GyoumuJouhouID =  " + data[0, 41];
                        cmd.ExecuteNonQuery();


                        //協力依頼書情報（KyouryokuIraisho）テーブル更新
                        //cmd.CommandText = "UPDATE KyouryokuIraisho SET " +
                        //    "KyouryokuChousaKijun = 1 " +
                        //    ",KyouryokuChousakijunbi = '" + CommonValue + "'" +
                        //    ",KyouryokuHoukokuSeigenDate = " + data[0, 7] + " " +
                        //    ",KyouryokuUpdateDate = SYSDATETIME()" +
                        //    ",KyouryokuUpdateUser ='" + UserInfos[0] + "' " +
                        //    ",KyouryokuUpdateProgram = '" + pgmName + methodName + "' " +
                        //    "WHERE MadoguchiID =" + MadoguchiID + " ";
                        cmd.CommandText = "UPDATE KyouryokuIraisho SET " +
                            "KyouryokuHoukokuSeigenDate = " + data[0, 7] + " " +
                            ",KyouryokuUpdateDate = SYSDATETIME()" +
                            ",KyouryokuUpdateUser = N'" + UserInfos[0] + "' " +
                            ",KyouryokuUpdateProgram = '" + pgmName + methodName + "' " +
                            "WHERE MadoguchiID =" + MadoguchiID + " ";

                        int updateCount = cmd.ExecuteNonQuery();

                        //update件数がなければ新規登録　
                        if (updateCount == 0)
                        {

                            //採番No（SaibanNo）を取得
                            var comboDt = new DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "SaibanNo+SaibanCountupNo AS SaibanNo " +
                              "FROM " + "M_Saiban " +
                              "WHERE SaibanMei = 'KyouryokuIraishoID' ";

                            //データ取得
                            var sda = new SqlDataAdapter(cmd);
                            sda.Fill(comboDt);
                            DataRow dr = comboDt.Rows[0];
                            int saibanKyouryokuNo = int.Parse(dr[0].ToString());

                            //採番No（SaibanNo）を更新
                            cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                                saibanKyouryokuNo + " WHERE SaibanMei = 'KyouryokuIraishoID' ";

                            cmd.ExecuteNonQuery();

                            //協力依頼書情報（KyouryokuIraisho）テーブル登録
                            cmd.CommandText = "INSERT INTO KyouryokuIraisho( " +
                                "KyouryokuIraishoID " +
                                ",MadoguchiID " +
                                ",KyouryokuChousaKijun " +
                                ",KyouryokuChousakijunbi " +
                                ",KyouryokuHoukokuSeigenDate " +
                                ",KyouryokuGyoumuKubun " +
                                ",KyouryokuIraiKubun " +
                                ",KyouryokuUtiawaseyouhi " +
                                ",KyouryokusakiHikiwatashi " +
                                ",KyouryokuJisshikeikakusho " +
                                ",KyourokuIraisakiTantoshaCD " +
                                ",KyouryokuCreateDate " +
                                ",KyouryokuCreateUser " +
                                ",KyouryokuCreateProgram " +
                                ",KyouryokuUpdateDate" +
                                ",KyouryokuUpdateUser " +
                                ",KyouryokuUpdateProgram " +
                                ",KyouryokuDeleteFlag)VALUES(" +
                                saibanKyouryokuNo +                    //採番
                                ",'" + MadoguchiID + "' " +               //窓口情報
                                ",'1' " +                              //1
                                ",N'" + CommonValue + "' " +            //M_COMMON_MASTER CHOUSAKIJUNBI_DEFAULT
                                "," + data[0, 7] + " " +  //調査担当者への締切日
                                ",NULL " +                             //空
                                ",NULL " +                             //空
                                ",'2' " +                              //2
                                ",'2' " +                              //2
                                ",'1' " +                              //1
                                ",NULL " +                             //NULL
                                ",SYSDATETIME() " +                             // 登録日時
                                ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                                ",'" + pgmName + methodName + "' " +            // 登録プログラム
                                ",SYSDATETIME() " +                             // 更新日時
                                ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                                ",'" + pgmName + methodName + "' " +            // 更新プログラム
                                ",0 " +                                         // 削除フラグ
                                ")";
                            cmd.ExecuteNonQuery();

                        }

                        // 意図がわかりかねるのでコメントアウト 2021/06/24
                        // 更新ボタン押下時にここのロジック通ることを確認した。
                        ////単品入力（TanpinNyuuryoku）テーブル更新
                        //cmd.CommandText = "UPDATE TanpinNyuuryoku SET " +
                        //    "TanpinGyoumuCD = 0 " +　//確定で0
                        //    "WHERE MadoguchiID =" + MadoguchiID + " ";

                        //updateCount = cmd.ExecuteNonQuery();
                        ////update件数がなければ新規登録
                        //if (updateCount == 0)
                        //{
                        //    //採番No（SaibanNo）を取得
                        //    var tanpinDt = new DataTable();
                        //    int TanpinNyuuryokuID = 0;
                        //    //SQL生成
                        //    cmd.CommandText = "SELECT " +
                        //      "SaibanNo+SaibanCountupNo AS SaibanNo " +
                        //      "FROM " + "M_Saiban " +
                        //      "WHERE SaibanMei = 'TanpinNyuuryokuID' ";

                        //    //データ取得
                        //    var tanpin_sda = new SqlDataAdapter(cmd);
                        //    tanpin_sda.Fill(tanpinDt);

                        //    DataRow tanpin_dr = tanpinDt.Rows[0];
                        //    TanpinNyuuryokuID = int.Parse(tanpin_dr[0].ToString());

                        //    //採番No（TanpinNyuuryokuID）を更新
                        //    cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                        //        TanpinNyuuryokuID + " WHERE SaibanMei = 'TanpinNyuuryokuID' ";

                        //    cmd.ExecuteNonQuery();
                        //    //単品入力（TanpinNyuuryoku）登録
                        //    cmd.CommandText = "INSERT INTO TanpinNyuuryoku(" +
                        //      "TanpinNyuuryokuID " +
                        //      ",MadoguchiID " +
                        //      ",TanpinGyoumuCD " +
                        //      ",TanpinDeleteFlag)VALUES(" +
                        //      TanpinNyuuryokuID +
                        //      ", '" + MadoguchiID + "' " +
                        //      ", " + "0" + " " + //確定で0
                        //      ",0)";
                        //    cmd.ExecuteNonQuery();
                        //}

                        ////採番No（SaibanNo）を取得
                        //var dt3 = new DataTable();
                        ////SQL生成
                        //cmd.CommandText = "SELECT " +
                        //  "SaibanNo+SaibanCountupNo AS SaibanNo " +
                        //  "FROM " + "M_Saiban " +
                        //  "WHERE SaibanMei = 'HistoryID' ";

                        ////データ取得
                        //var sda3 = new SqlDataAdapter(cmd);
                        //sda3.Fill(dt3);

                        //DataRow dr3 = dt3.Rows[0];
                        //int saibanHistoryNo = int.Parse(dr3[0].ToString());

                        ////採番No（SaibanNo）を更新
                        //cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                        //    saibanHistoryNo + " WHERE SaibanMei = 'HistoryID' ";

                        //cmd.ExecuteNonQuery();

                        //履歴登録
                        //cmd.CommandText = "INSERT INTO T_HISTORY(" +
                        //    "H_DATE_KEY " +
                        //    ",H_NO_KEY " +
                        //    ",H_OPERATE_DT " +
                        //    ",H_OPERATE_USER_ID " +
                        //    ",H_OPERATE_USER_MEI " +
                        //    ",H_OPERATE_USER_BUSHO_CD " +
                        //    ",H_OPERATE_USER_BUSHO_MEI " +
                        //    ",H_OPERATE_NAIYO " +
                        //    ",H_ProgramName " +
                        //    ",MadoguchiID " +
                        //    ",HistoryBeforeTantoubushoCD " +
                        //    ",HistoryBeforeTantoushaCD " +
                        //    ",HistoryAfterTantoubushoCD " +
                        //    ",HistoryAfterTantoushaCD " +
                        //    ")VALUES(" +
                        //    "SYSDATETIME() " + 
                        //    ", " + saibanHistoryNo + " " +
                        //    ",SYSDATETIME() " +
                        //    ",'" + UserInfos[0] + "' " +
                        //    ",'" + UserInfos[1] + "' " +
                        //    ",'" + UserInfos[2] + "' " +
                        //    ",'" + UserInfos[3] + "' " +
                        //    ",'調査概要を更新しました ID:" + MadoguchiID + " Garoon連携区分:" + data[0, 44] + "' " +
                        //    ",'" + pgmName + methodName + "' " +
                        //    "," + MadoguchiID + " " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ")";
                        //cmd.ExecuteNonQuery();s

                        // 案件情報IDが異なる⇒受託番号変更を行われた
                        //if (AnkenJouhouID != data[0, 41])
                        // 受託番号変更あり 
                        if (data[0, 51] != null && data[0, 51] == "1")
                        {
                            // 単価契約の取得
                            var TankaKeiyakuDt = new DataTable();
                            int TankaKeiyakuID = 0;
                            //SQL生成
                            cmd.CommandText = "SELECT TankaKeiyakuID FROM TankaKeiyaku"
                                            + " LEFT JOIN AnkenJouhou ON TankaKeiyaku.TankakeiyakuJutakuBangou = AnkenJouhou.AnkenJutakuBangou"
                                            + " WHERE AnkenJouhou.AnkenJouhouID = " + data[0, 41]
                                            + " ORDER BY TankaKeiyakuID DESC ";

                            //データ取得
                            var TankaKeiyaku_sda = new SqlDataAdapter(cmd);
                            TankaKeiyaku_sda.Fill(TankaKeiyakuDt);
                            if (TankaKeiyakuDt.Rows.Count > 0)
                            {
                                TankaKeiyakuID = int.Parse(TankaKeiyakuDt.Rows[0][0].ToString());
                            }
                            if (TankaKeiyakuID != 0)
                            {
                                cmd.CommandText = "UPDATE TanpinNyuuryoku set " +
                                  "TanpinGyoumuCD = " + TankaKeiyakuID +
                                  ",TanpinHachuubusho = N'" + data[0, 45] + "' " +
                                  ",TanpinYakushoku = N'" + data[0, 46] + "' " +
                                  ",TanpinHachuuTantousha = N'" + data[0, 47] + "' " +
                                  ",TanpinTel = N'" + data[0, 48] + "' " +
                                  ",TanpinFax = N'" + data[0, 49] + "' " +
                                  ",TanpinMail = N'" + data[0, 50] + "' " +
                                  ",TanpinUpdateDate = SYSDATETIME() " +
                                  ",TanpinUpdateUser = N'" + UserInfos[0] + "' " +
                                  ",TanpinUpdateProgram = '" + pgmName + methodName + "' " +
                                  "WHERE MadoguchiID = '" + MadoguchiID + "' ";
                                cmd.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();

                        Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "調査概要を更新しました ID:" + MadoguchiID + " Garoon連携区分:" + data[0, 44], pgmName + methodName, MadoguchiID);

                        transaction = sqlconn.BeginTransaction();
                        cmd.Transaction = transaction;

                        //応援受付（OuenUketsuke）テーブル更新
                        String kanriNo = "";
                        if (!String.IsNullOrEmpty(data[0, 22]))
                        {
                            kanriNo = data[0, 22];
                        }

                        String ouenJoukyo = "0";
                        // 支→本の場合、1とする
                        if(data[0, 21] != "0")
                        {
                            ouenJoukyo = "1";
                        }

                        cmd.CommandText = "UPDATE OuenUketsuke SET " +
                            "OuenKanriNo = '" + kanriNo + "' " +
                            ", OuenUpdateDate = SYSDATETIME() " +
                            ", OuenUpdateUser = N'" + UserInfos[0] + "' " +
                            ", OuenUpdateProgram = '" + pgmName + methodName + "' " +
                            ", OuenJoukyou = CASE OuenJoukyou WHEN 2 THEN 2 ELSE " + ouenJoukyo + " END " +
                            " WHERE MadoguchiID = " + MadoguchiID + " ";
                        //Clipboard.SetText(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        //実施区分が3の中止の場合
                        if ("3".Equals(data[0, 11]))
                        {
                            //MadoguchiL1ChousaShinchoku が 6ではないデータを更新
                            //cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                            //"MadoguchiL1ChousaShinchoku = 6 " +
                            //",MadoguchiL1ChousaKakunin = 1 " +
                            //" WHERE MadoguchiL1ChousaShinchoku != 6 ";
                            cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                            "MadoguchiL1ChousaShinchoku = 80 " +
                            ",MadoguchiL1ChousaKakunin = 1 " +
                            ",MadoguchiL1AsteriaKoushinFlag = 1 " +
                            ",MadoguchiL1UpdateDate = SYSDATETIME() " +
                            ",MadoguchiL1UpdateUser = N'" + UserInfos[0] + "' " +
                            ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
                            " WHERE MadoguchiL1ChousaShinchoku != 80 AND MadoguchiID = " + MadoguchiID;

                            cmd.ExecuteNonQuery();

                            //ChousaShinchokuJoukyou が 6ではないデータを更新
                            //cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                            //"ChousaShinchokuJoukyou = 6 " +
                            //",ChousaHoukokuzumi = 1 " +
                            //" WHERE ChousaShinchokuJoukyou != 6 ";
                            cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                            "ChousaShinchokuJoukyou = 80 " +
                            ",ChousaHoukokuzumi = 1 " +
                            ",ChousaUpdateDate = SYSDATETIME() " +
                            ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +
                            ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                            " WHERE ChousaShinchokuJoukyou != 80 AND MadoguchiID = " + MadoguchiID;

                            cmd.ExecuteNonQuery();

                        }
                        else
                        {
                            // 報告済みにチェックが入っているときに二次検証済みとして更新する
                            // 1:報告済み
                            if ("1".Equals(data[0, 34]))
                            {
                                cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                                "MadoguchiL1ChousaShinchoku = 70 " +
                                ",MadoguchiL1ChousaKakunin = 1 " +
                                ",MadoguchiL1UpdateDate = SYSDATETIME() " +
                                ",MadoguchiL1UpdateUser = N'" + UserInfos[0] + "' " +
                                ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
                                " WHERE MadoguchiL1ChousaShinchoku != 80 AND MadoguchiID = " + MadoguchiID;
                                //" WHERE MadoguchiID = " + MadoguchiID;

                                cmd.ExecuteNonQuery();

                                //ChousaShinchokuJoukyou が 6ではないデータを更新
                                cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                                "ChousaShinchokuJoukyou = 70 " +
                                ",ChousaHoukokuzumi = 1 " +
                                ",ChousaUpdateDate = SYSDATETIME() " +
                                ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +
                                ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                                " WHERE ChousaShinchokuJoukyou != 80 AND MadoguchiID = " + MadoguchiID;
                                //" WHERE MadoguchiID = " + MadoguchiID;

                                cmd.ExecuteNonQuery();

                            }
                            // 0:報告済みでない
                            else if ("0".Equals(data[0, 34]))
                            {
                               // 1180 仮対処でコメントアウト
                               // cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                               // "MadoguchiL1ChousaShinchoku = " + shinchoku + " " + // 親と同じ進捗に合わせる
                               // ",MadoguchiL1UpdateDate = SYSDATETIME() " +
                               // ",MadoguchiL1UpdateUser = '" + UserInfos[0] + "' " +
                               // ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
                               // " WHERE MadoguchiID = " + MadoguchiID;

                               // cmd.ExecuteNonQuery();

                               // cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                               //"ChousaHoukokuzumi = 0 " +
                               //",ChousaShinchokuJoukyou = " + shinchoku + " " + // 親と同じ進捗に合わせる
                               //",ChousaUpdateDate = SYSDATETIME() " +
                               //",ChousaUpdateUser = '" + UserInfos[0] + "' " +
                               //",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                               //" WHERE MadoguchiID = " + MadoguchiID;

                               // cmd.ExecuteNonQuery();
                            }
                        }

                        // 親情報に変更があったので、担当部所の親データを更新する
                        if(parentUpdateFlg == true)
                        {
                            string updateDateStr = DateTime.Now.ToString();

                            cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                            " MadoguchiL1TourokuNendo = '" + data[0,2] + "' " +
                            ",MadoguchiL1UketsukeBangou = N'" + data[0,25] + "' " +
                            ",MadoguchiL1UketsukeBangouEdaban = N'" + data[0, 26] + "' " +
                            ",MadoguchiL1TokuchoBangou = N'" + data[0, 25] + "-" + data[0, 26] + "' " +
                            ",MadoguchiL1HachushaMei = N'" + data[0, 27] + "' " +
                            ",MadoguchiL1ShukeihyoLink = N'" + data[0, 37] + "' " +
                            ",MadoguchiL1UpdateDate = '" + updateDateStr + "' " +
                            ",MadoguchiL1UpdateUser = N'" + UserInfos[0] + "' " +
                            ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
                            " WHERE MadoguchiID  = '" + MadoguchiID + "' ";

                            cmd.ExecuteNonQuery();

                            Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "親情報に変更があったため、担当部所を更新しました。 窓口ID = " + MadoguchiID, "MadoguchiUpdate_SQL", MadoguchiID);

                        }

                        // 窓口部所が変更された場合、支部応援をGaroonTsuikaAtesakiに追加する
                        // 変更前の窓口部所と画面の窓口部所
                        if (MadoguchiTantoushaBushoCD != data[0,15])
                        {
                            // GaroonTsuikaAtesakiに窓口部所に一致する支部応援のユーザーを追加する
                            var garoonDt = new DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "Mst_Chousain.KojinCD " +
                              ",Mst_Chousain.ChousainMei " +
                              ",Mst_Chousain.GyoumuBushoCD " +
                              ",mb.BushokanriboKamei " +
                              "FROM Mst_Shibuouen INNER JOIN Mst_Chousain ON " +
                              "Mst_Chousain.KojinCD = Mst_Shibuouen.ShibuouenKojinCD " +
                              //"AND Mst_Shibuouen.ShibuouenDeleteFlag = 0 " +
                              //"AND Mst_Chousain.RetireFLG = 0 " +
                              "AND Mst_Chousain.GyoumuBushoCD ='" + data[0,15] + "' " +
                              "INNER JOIN Mst_Busho mb ON mb.GyoumuBushoCD = Mst_Chousain.GyoumuBushoCD ";
                            //データ取得
                            var garronSda = new SqlDataAdapter(cmd);
                            garronSda.Fill(garoonDt);

                            string KojinCD = "";
                            string ChousainMei = "";
                            string GyoumuBushoCD = "";
                            string BushoMei = "";

                            if (garoonDt != null && garoonDt.Rows.Count > 0)
                            {
                                for (int i = 0; i < garoonDt.Rows.Count; i++)
                                {
                                    KojinCD = garoonDt.Rows[i][0].ToString();
                                    ChousainMei = garoonDt.Rows[i][1].ToString();
                                    GyoumuBushoCD = garoonDt.Rows[i][2].ToString();
                                    BushoMei = garoonDt.Rows[i][3].ToString();

                                    // 窓口部所に一致する支部応援で既に登録済みかもしれないので、存在チェック
                                    string where = "GaroonTsuikaAtesakiMadoguchiID = '" + data[0,1] + "' " +
                                                   "AND GaroonTsuikaAtesakiBushoCD = '" + GyoumuBushoCD + "' " +
                                                   "AND GaroonTsuikaAtesakiTantoushaCD = '" + KojinCD + "' " +
                                                   " AND GaroonTsuikaAtesakiDeleteFlag <> 1";

                                    var tmpdt = getData("GaroonTsuikaAtesakiTantoushaCD", "GaroonTsuikaAtesakiTantoushaCD", "GaroonTsuikaAtesaki with(nolock) ", where);
                                    // データ件数が0件なら登録
                                    if (tmpdt != null && tmpdt.Rows.Count == 0)
                                    {

                                        // GaroonTsuikaAtesakiに登録
                                        cmd.CommandText = "INSERT INTO GaroonTsuikaAtesaki ( " +
                                        " GaroonTsuikaAtesakiID " +
                                        ",GaroonTsuikaAtesakiMadoguchiID " +
                                        ",GaroonTsuikaAtesakiBushoCD " +
                                        ",GaroonTsuikaAtesakiBusho " +
                                        ",GaroonTsuikaAtesakiTantoushaCD " +
                                        ",GaroonTsuikaAtesakiTantousha " +
                                        ",GaroonTsuikaAtesakiCreateDate " +
                                        ",GaroonTsuikaAtesakiCreateUser " +
                                        ",GaroonTsuikaAtesakiCreateProgram " +
                                        ",GaroonTsuikaAtesakiUpdateDate " +
                                        ",GaroonTsuikaAtesakiUpdateUser " +
                                        ",GaroonTsuikaAtesakiUpdateProgram " +
                                        ",GaroonTsuikaAtesakiDeleteFlag " +
                                        ") VALUES (" +
                                        "'" + getSaiban("GaroonTsuikaAtesakiID") + "' " + // GaroonTsuikaAtesakiID
                                        ",'" + data[0, 1] + "' " +          // GaroonTsuikaAtesakiMadoguchiID
                                        ",'" + GyoumuBushoCD + "' " +              // GaroonTsuikaAtesakiBushoCD
                                        ",N'" + BushoMei + "' " +                   // GaroonTsuikaAtesakiBusho
                                        ",'" + KojinCD + "' " +                    // GaroonTsuikaAtesakiTantoushaCD
                                        ",N'" + ChousainMei + "' " +                // GaroonTsuikaAtesakiTantousha
                                        ",SYSDATETIME() " +                             // 登録日時
                                        ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                                        ",'" + pgmName + methodName + "' " +            // 登録プログラム
                                        ",SYSDATETIME() " +                             // 更新日時
                                        ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                                        ",'" + pgmName + methodName + "' " +            // 更新プログラム
                                        ",0 " +                                         // 削除フラグ
                                        ") ";

                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }
                        }

                        string nowMadoguchiTourokubi = "";
                        string nowMadoguchiJutakuBangou = "";
                        string nowMadoguchiJutakuBangouEdaban = "";
                        string nowMadoguchiUketsukeBangou = "";
                        string nowMadoguchiUketsukeBangouEdaban = "";

                        // 898 登録日、受託番号変更時にAsteriaFlgをONするする対応
                        // 更新した登録日、受託番号、枝番を取得
                        var dtNow = new DataTable();
                        cmd.CommandText = "SELECT MadoguchiTourokubi,MadoguchiJutakuBangou,MadoguchiJutakuBangouEdaban,MadoguchiUketsukeBangou,MadoguchiUketsukeBangouEdaban " +
                            "FROM MadoguchiJouhou " +
                            "WHERE MadoguchiID = " + MadoguchiID + "";
                        //データ取得
                        var nowSda = new SqlDataAdapter(cmd);
                        nowSda.Fill(dtNow);

                        if (dtNow != null && dtNow.Rows.Count > 0)
                        {
                            nowMadoguchiTourokubi = dtNow.Rows[0][0].ToString();
                            nowMadoguchiJutakuBangou = dtNow.Rows[0][1].ToString();
                            nowMadoguchiJutakuBangouEdaban = dtNow.Rows[0][2].ToString();
                            nowMadoguchiUketsukeBangou = dtNow.Rows[0][3].ToString();
                            nowMadoguchiUketsukeBangouEdaban = dtNow.Rows[0][4].ToString();

                            // 登録日,受託番号,特調番号が違う場合
                            if (MadoguchiTourokubi != nowMadoguchiTourokubi 
                                || MadoguchiJutakuBangou != nowMadoguchiJutakuBangou || MadoguchiJutakuBangouEdaban != nowMadoguchiJutakuBangouEdaban
                                || MadoguchiUketsukeBangou != nowMadoguchiUketsukeBangou || MadoguchiUketsukeBangouEdaban != nowMadoguchiUketsukeBangouEdaban
                                )
                            {
                                cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                                    " MadoguchiL1MitsumoriFrom = N'" + nowMadoguchiTourokubi + "' " +
                                    ",MadoguchiL1UketsukeBangou = N'" + nowMadoguchiUketsukeBangou + "' " +
                                    ",MadoguchiL1UketsukeBangouEdaban = N'" + nowMadoguchiUketsukeBangouEdaban + "' " +
                                    ",MadoguchiL1TokuchoBangou = N'" + nowMadoguchiUketsukeBangou + "-" + nowMadoguchiUketsukeBangouEdaban + "' " +
                                    ",MadoguchiL1UpdateDate = '" + DateTime.Now.ToString() + "' " +
                                    ",MadoguchiL1UpdateUser = N'" + UserInfos[0] + "' " +
                                    ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
                                    ",MadoguchiL1TokuchoHaitaFlag = 1 " +
                                    ",MadoguchiL1AsteriaKoushinFlag = 1 " + // Asteriaフラグも立てておく
                                    " WHERE MadoguchiID  = '" + MadoguchiID + "' ";

                                cmd.ExecuteNonQuery();

                            }
                        }
                    }
                    // I20102:データを更新しました。
                    mes += GetMessage("I20102", "");
                    /*
                    //新規Garoon追加宛先処理（作成中）
                    //Garoon追加宛先　自動追加処理
                    string[,] TsuikaAtesaki = new string[4,100];
                    //管理技術者の取得
                    var KanriDT = new DataTable();
                    cmd.CommandText = "SELECT " +
                        "Mst_Chousain.GyoumuBushoCD, Mst_Busho.BushokanriboKamei, MadoguchiKanriGijutsusha, ChousainMei " +
                        "FROM MadoguchiJouhou " +
                        "LEFT JOIN Mst_Chousain ON MadoguchiKanriGijutsusha = ChousainCD " +
                        "LEFT JOIN Mst_Busho ON Mst_Busho.GyoumuBushoCD = Mst_Chousain.GyoumuBushoCD " +
                        "WHERE MadoguchiID = '" + MadoguchiID + "' ";

                    if (KanriDT != null && KanriDT.Rows.Count > 0)
                    {
                        TsuikaAtesaki[0,0] = KanriDT.Rows[0][0].ToString();
                        TsuikaAtesaki[0,1] = KanriDT.Rows[0][1].ToString();
                        TsuikaAtesaki[0,2] = KanriDT.Rows[0][2].ToString();
                        TsuikaAtesaki[0,3] = KanriDT.Rows[0][3].ToString();
                    }

                    //調査担当者の部所応援を取得
                    var ShibuOuenDT = new DataTable();
                    cmd.CommandText = "SELECT " +
                        "Mst_Chousain.GyoumuBushoCD, Mst_Busho.BushokanriboKamei, ChousainCD, ChousainMei " +
                        "FROM MadoguchiJouhouMadoguchiL1Chou " +
                        "LEFT JOIN Mst_Chousain ON Mst_Chousain.GyoumuBushoCD = MadoguchiL1ChousaBushoCD AND Mst_Chousain.RetireFLG <> 1 " +
                        "INNER JOIN Mst_Shibuouen ON ShibuouenKojinCD = ChousainCD AND Mst_Shibuouen.ShibuouenDeleteFlag <> 1 " +
                        "LEFT JOIN Mst_Busho ON Mst_Busho.GyoumuBushoCD = Mst_Chousain.GyoumuBushoCD " +
                        "WHERE MadoguchiJouhouMadoguchiL1Chou.MadoguchiID = '" + MadoguchiID + "' ";

                    if (ShibuOuenDT != null && ShibuOuenDT.Rows.Count > 0)
                    {
                        for (int i = 0; i < ShibuOuenDT.Rows.Count; i++)
                        {
                            TsuikaAtesaki[i + 1, 0] = ShibuOuenDT.Rows[i][0].ToString();
                            TsuikaAtesaki[i + 1, 1] = ShibuOuenDT.Rows[i][1].ToString();
                            TsuikaAtesaki[i + 1, 2] = ShibuOuenDT.Rows[i][2].ToString();
                            TsuikaAtesaki[i + 1, 3] = ShibuOuenDT.Rows[i][3].ToString();
                        }
                    }
                    //追加宛先の重複削除

                    //Garoon追加宛先テーブル　存在確認

                        //登録済みの場合は、変更なし

                        //未登録の場合は登録
                    */
                }
                // 担当部所
                else if (tab == 2)
                {
                    // メッセージ表示フラグ
                    //調査員更新メッセージフラグ
                    int messageFlg = 0;
                    //調査員削除メッセージフラグ
                    int messageFlg1 = 0;
                    //調査品目削除メッセージフラグ
                    int messageFlg2 = 0;
                    //調査品目更新・削除なしメッセージフラグ
                    int messageFlg3 = 0;
                    //調査品目更新メッセージフラグ
                    int messageFlg4 = 0;
                    //締切日変更メッセージフラグ
                    int messageFlg5 = 0;
                    // 支部備考メッセージフラグ
                    int shibuMessageFlg = 0;

                    string bikou = "○";

                    if ("Tokumei".Equals(gamenMode))
                    {
                        bikou = "◎";
                    }

                    for (int i = 0; i < data.GetLength(0); i++)
                    {
                        // データが無い場合、処理を抜ける
                        if (data[i, 0] == null)
                        {
                            break;
                        }

                        //現在の窓口担当者データを取得（締切日、進捗状況、引渡フラグ）
                        cmd.CommandText = "SELECT MadoguchiL1ChousaShimekiribi, MadoguchiL1ChousaShinchoku, MadoguchiL1ChousaKakunin " +
                                            " FROM MadoguchiJouhouMadoguchiL1Chou WHERE MadoguchiID = " + MadoguchiID + " AND MadoguchiL1ChousaCD = " + data[i, 0];

                        var dt = new DataTable();
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(dt);

                        //データが存在する場合
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            //変更した項目がある場合
                            // 0:締切日
                            // 1:進捗
                            // 2:引渡し完了
                            if (!dt.Rows[0][0].ToString().Equals(data[i, 3]) || !dt.Rows[0][1].ToString().Equals(data[i, 4]) || !dt.Rows[0][2].ToString().Equals(data[i, 5]))
                            {
                                //変更内容を更新
                                cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET ";
                                if (data[i, 3] != null && data[i, 3] != "")
                                {
                                    cmd.CommandText += "MadoguchiL1ChousaShimekiribi = '" + data[i, 3] + "' ";
                                }
                                else
                                {
                                    cmd.CommandText += "MadoguchiL1ChousaShimekiribi = NULL ";
                                }
                                cmd.CommandText += ",MadoguchiL1ChousaShinchoku = '" + data[i, 4] + "' " +
                                                    ",MadoguchiL1ChousaKakunin = '" + data[i, 5] + "' " +
                                                    ",MadoguchiL1AsteriaKoushinFlag = 1 " +
                                                    ",MadoguchiL1UpdateDate = SYSDATETIME() " +
                                                    ",MadoguchiL1UpdateUser = N'" + UserInfos[0] + "' " +
                                                    ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
                                                    " WHERE MadoguchiID = " + MadoguchiID + " AND MadoguchiL1ChousaCD = " + data[i, 0];

                                Console.WriteLine(cmd.CommandText);
                                cmd.ExecuteNonQuery();
                                messageFlg = 1;
                                Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "調査員の項目を更新しました。  窓口ID = " + MadoguchiID + " 部所: " + data[i, 1] + " 調査員:" + data[i, 2], "MadoguchiUpdate_SQL", MadoguchiID);

                            }
                            //変更項目がない場合
                            else
                            {
                                //調査員のデータへの更新がないので↓のメッセージは使わない

                                // I20205:調査品目明細画面への更新はありませんでした。
                                //mes += GetMessage("I20205", "");
                            }

                            //担当者進捗状況が変更された場合
                            if (!dt.Rows[0][1].ToString().Equals(data[i, 4]))
                            {

                                if (GetCommonValue1("CHOUSA_CHRNGE").Equals("1"))
                                {
                                    // 719 ②調査中の進捗状況が増えているので、進捗状況を変更した際、調査開始、見積中、集計中と調査品目の進捗状況を変更していかないといけない
                                    // 現行だとChousaShinchokuJoukyou = 70を見ているが、その条件は見ずに、担当部所の進捗で更新する
                                    cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                                                        "ChousaShinchokuJoukyou = '" + data[i, 4] + "' " +
                                                        ",ChousaUpdateDate = SYSDATETIME() " +
                                                        ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +
                                                        ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                                                        " WHERE HinmokuRyakuBushoCD = '" + data[i, 1] + "' AND HinmokuChousainCD = '" + data[i, 2] + "' " +
                                                        "AND MadoguchiID = '" + MadoguchiID + "' ";

                                    ////70：二次検済　に変更した場合
                                    //if (data[i, 4].Equals("70"))
                                    //{
                                    //    cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                                    //                        "ChousaShinchokuJoukyou = '" + data[i, 4] + "' " +
                                    //                        ",ChousaUpdateDate = SYSDATETIME() " +
                                    //                        ",ChousaUpdateUser = '" + UserInfos[0] + "' " +
                                    //                        ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                                    //                        " WHERE HinmokuRyakuBushoCD = '" + data[i, 1] + "' AND HinmokuChousainCD = '" + data[i, 2] + "' " +
                                    //                        "AND MadoguchiID = '" + MadoguchiID + "' ";
                                    //}
                                    ////70：二次検済　以外に変更した場合は、70：二次検済のデータを更新対象とする
                                    //else
                                    //{
                                    //    cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                                    //                        "ChousaShinchokuJoukyou = '" + data[i, 4] + "' " +
                                    //                        ",ChousaUpdateDate = SYSDATETIME() " +
                                    //                        ",ChousaUpdateUser = '" + UserInfos[0] + "' " +
                                    //                        ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                                    //                        " WHERE HinmokuRyakuBushoCD = '" + data[i, 1] + "' AND HinmokuChousainCD = '" + data[i, 2] + "' AND ChousaShinchokuJoukyou = 70 " +
                                    //                        "AND MadoguchiID = '" + MadoguchiID + "' ";
                                    //}

                                    Console.WriteLine(cmd.CommandText);
                                    var result = cmd.ExecuteNonQuery();

                                    if (result > 0)
                                    {
                                        Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "調査品目を更新しました。  窓口ID = " + MadoguchiID + " 部所: " + data[i, 1] + " 調査員:" + data[i, 2], "MadoguchiUpdate_SQL", MadoguchiID);
                                    }

                                    //cmd.CommandText = "UPDATE ShibuBikou SET " +
                                    //                    "ShibuBikouChousaBusho = '" + bikou + "' " +
                                    //                    ",ShibuBikouUpdateDate = SYSDATETIME() " +
                                    //                    ",ShibuBikouUpdateUser = '" + UserInfos[0] + "' " +
                                    //                    ",ShibuBikouUpdateProgram = '" + pgmName + methodName + "' " +
                                    //                    " WHERE ShibuBikouBushoKanriboBushoCD = '" + data[i, 1] + "' AND MadoguchiID = '" + MadoguchiID + "' ";

                                    //Console.WriteLine(cmd.CommandText);
                                    //result = cmd.ExecuteNonQuery();

                                    //if (result <= 0)
                                    //{
                                    //    cmd.CommandText = "INSERT INTO ShibuBikou ( " +
                                    //                        "MadoguchiID " +
                                    //                        ",ShibuBikouID " +
                                    //                        ",ShibuBikouBushoKanriboBushoCD " +
                                    //                        ",ShibuBikouKanriNo " +
                                    //                        ",ShibuBikouChousaBusho " +
                                    //                        ",ShibuBikou " +
                                    //                        ",ShibuBikouRyakumei " +
                                    //                        ",ShibuBikouCreateDate " +
                                    //                        ",ShibuBikouCreateUser " +
                                    //                        ",ShibuBikouCreateProgram " +
                                    //                        ",ShibuBikouUpdateDate " +
                                    //                        ",ShibuBikouUpdateUser " +
                                    //                        ",ShibuBikouUpdateProgram " +
                                    //                        ",ShinDeleteFlag " +
                                    //                        ",ShibuBushokanriboKameiRakuOld " +
                                    //                        ",ShibuBushoKanriboShibuMeiOld " +
                                    //                        ") VALUES  ( " +
                                    //                        MadoguchiID +
                                    //                        "," + getSaiban("ShibuBikouID") +
                                    //                        ", " + data[i, 1] +
                                    //                        ",NULL " +
                                    //                        ",'" + bikou + "' " +
                                    //                        ",NULL " +
                                    //                        ", " + data[i, 2] +
                                    //                        ",SYSDATETIME() " +                             // 登録日時
                                    //                        ",'" + UserInfos[0] + "' " +                    // 登録ユーザ
                                    //                        ",'" + pgmName + methodName + "' " +            // 登録プログラム
                                    //                        ",SYSDATETIME() " +                             // 更新日時
                                    //                        ",'" + UserInfos[0] + "' " +                    // 更新ユーザ
                                    //                        ",'" + pgmName + methodName + "' " +            // 更新プログラム
                                    //                        ",0 " +                                         // 削除フラグ
                                    //                        ",NULL " +
                                    //                        ",NULL " +
                                    //                        " ) ";

                                    //    Console.WriteLine(cmd.CommandText);
                                    //    result = cmd.ExecuteNonQuery();
                                    //    Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "支部備考を登録しました。  窓口ID = " + MadoguchiID + " 部所: " + data[i, 1] + " 調査員:" + data[i, 2], "MadoguchiUpdate_SQL", MadoguchiID);

                                    //}
                                    //else
                                    //{
                                    //    Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "支部備考を更新しました。  窓口ID = " + MadoguchiID + " 部所: " + data[i, 1] + " 調査員:" + data[i, 2], "MadoguchiUpdate_SQL", MadoguchiID);
                                    //}

                                    DataTable shibuDT = new DataTable();
                                    shibuDT = getData("ShibuBikouChousaBusho", "ShibuBikouChousaBusho", "ShibuBikou with(nolock) ", "ShibuBikouBushoKanriboBushoCD = '" + data[i, 1] + "' AND MadoguchiID = '" + MadoguchiID + "'");
                                    string chousaBusho = "";

                                    if (shibuDT != null && shibuDT.Rows.Count > 0)
                                    {
                                        chousaBusho = shibuDT.Rows[0][0].ToString();

                                        // ShibuBikouChousaBushoに変更があるなら更新する
                                        if (chousaBusho != bikou)
                                        {
                                            cmd.CommandText = "UPDATE ShibuBikou SET " +
                                                        "ShibuBikouChousaBusho = N'" + bikou + "' " +
                                                        ",ShibuBikouUpdateDate = SYSDATETIME()  " +
                                                        ",ShibuBikouUpdateUser = N'" + UserInfos[0] + "' " +
                                                        ",ShibuBikouUpdateProgram = '" + pgmName + methodName + "' " +
                                                        " WHERE ShibuBikouBushoKanriboBushoCD = '" + data[i, 1] + "' AND MadoguchiID = '" + MadoguchiID + "' ";

                                            Console.WriteLine(cmd.CommandText);
                                            result = cmd.ExecuteNonQuery();
                                            shibuMessageFlg = 1;
                                            Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "支部備考を更新しました。  窓口ID = " + MadoguchiID + " 部所: " + data[i, 1] + " 調査員:" + data[i, 2], "MadoguchiUpdate_SQL", MadoguchiID);
                                        }
                                    }
                                    else
                                    {
                                        cmd.CommandText = "INSERT INTO ShibuBikou ( " +
                                                        "MadoguchiID " +
                                                        ",ShibuBikouID " +
                                                        ",ShibuBikouBushoKanriboBushoCD " +
                                                        ",ShibuBikouKanriNo " +
                                                        ",ShibuBikouChousaBusho " +
                                                        ",ShibuBikou " +
                                                        ",ShibuBikouRyakumei " +
                                                        ",ShibuBikouCreateDate " +
                                                        ",ShibuBikouCreateUser " +
                                                        ",ShibuBikouCreateProgram " +
                                                        ",ShibuBikouUpdateDate " +
                                                        ",ShibuBikouUpdateUser " +
                                                        ",ShibuBikouUpdateProgram " +
                                                        ",ShinDeleteFlag " +
                                                        ",ShibuBushokanriboKameiRakuOld " +
                                                        ",ShibuBushoKanriboShibuMeiOld " +
                                                        ") VALUES  ( " +
                                                        MadoguchiID +
                                                        "," + getSaiban("ShibuBikouID") +
                                                        ", " + data[i, 1] +
                                                        ",NULL " +
                                                        ",N'" + bikou + "' " +
                                                        ",NULL " +
                                                        //", " + data[i, 2] +
                                                        ", (SELECT BushokanriboKameiRaku FROM MST_Busho WHERE GyoumuBushoCD = '" + data[i, 1] + "')" +
                                                        ",SYSDATETIME() " +                             // 登録日時
                                                        ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                                                        ",'" + pgmName + methodName + "' " +            // 登録プログラム
                                                        ",SYSDATETIME() " +                             // 更新日時
                                                        ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                                                        ",'" + pgmName + methodName + "' " +            // 更新プログラム
                                                        ",0 " +                                         // 削除フラグ
                                                        ",NULL " +
                                                        ",NULL " +
                                                        " ) ";

                                        Console.WriteLine(cmd.CommandText);
                                        result = cmd.ExecuteNonQuery();
                                        shibuMessageFlg = 1;
                                        Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "支部備考を登録しました。  窓口ID = " + MadoguchiID + " 部所: " + data[i, 1] + " 調査員:" + data[i, 2], "MadoguchiUpdate_SQL", MadoguchiID);
                                    }


                                    //// I20207:支部備考を更新しました。
                                    //mes += Environment.NewLine + GetMessage("I20207", "");
                                }
                            }
                            //締切日が変更された場合
                            //else if (!dt.Rows[0][0].ToString().Equals(data[i, 3]))
                            if (!dt.Rows[0][0].ToString().Equals(data[i, 3]))
                            {

                                if (GetCommonValue1("CHOUSA_CHRNGE").Equals("1"))
                                {

                                    //締切日一括更新解除（■チェック有：調査中の調査品目の締切日が変更されます。　■チェック無；締切日が一括更新されます。）
                                    if (!data[i, 6].Equals("1"))
                                    {
                                        cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                                                            "ChousaHinmokuShimekiribi = '" + data[i, 3] + "' " +
                                                            ",ChousaUpdateDate = SYSDATETIME() " +
                                                            ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +
                                                            ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                                                            " WHERE HinmokuRyakuBushoCD = '" + data[i, 1] + "' AND HinmokuChousainCD = '" + data[i, 2] + "' " +
                                                            "AND MadoguchiID = '" + MadoguchiID + "' ";
                                    }
                                    else
                                    {
                                        //40：集計中（調査中）のみ更新
                                        // 20:調査開始 のみ更新
                                        // 20:調査開始 30:見積中 40:集計中
                                        cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                                                            "ChousaHinmokuShimekiribi = '" + data[i, 3] + "' " +
                                                            ",ChousaUpdateDate = SYSDATETIME() " +
                                                            ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +
                                                            ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                                                            //" WHERE HinmokuRyakuBushoCD = '" + data[i, 1] + "' AND HinmokuChousainCD = '" + data[i, 2] + "' AND ChousaShinchokuJoukyou = 40 " +
                                                            //" WHERE HinmokuRyakuBushoCD = '" + data[i, 1] + "' AND HinmokuChousainCD = '" + data[i, 2] + "' AND ChousaShinchokuJoukyou = 20 " + // 20:調査開始
                                                            " WHERE HinmokuRyakuBushoCD = '" + data[i, 1] + "' AND HinmokuChousainCD = '" + data[i, 2] + "' AND ChousaShinchokuJoukyou in(20,30,40) " + // 20:調査開始 30:見積中 40:集計中
                                                            "AND MadoguchiID = '" + MadoguchiID + "' ";
                                    }
                                    Console.WriteLine(cmd.CommandText);
                                    var result = cmd.ExecuteNonQuery();

                                    //70：二次検済
                                    if (data[i, 4].Equals("70"))
                                    {
                                        cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                                                            "ChousaShinchokuJoukyou = '" + data[i, 4] + "' " +
                                                            ",ChousaUpdateDate = SYSDATETIME() " +
                                                            ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +
                                                            ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                                                            " WHERE HinmokuRyakuBushoCD = '" + data[i, 1] + "' AND HinmokuChousainCD = '" + data[i, 2] + "' " +
                                                            "AND MadoguchiID = '" + MadoguchiID + "' ";
                                    }
                                    else
                                    {
                                        cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                                                            "ChousaShinchokuJoukyou = '" + data[i, 4] + "' " +
                                                            ",ChousaUpdateDate = SYSDATETIME() " +
                                                            ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +
                                                            ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                                                            " WHERE HinmokuRyakuBushoCD = '" + data[i, 1] + "' AND HinmokuChousainCD = '" + data[i, 2] + "' AND ChousaShinchokuJoukyou = 70 " +
                                                            "AND MadoguchiID = '" + MadoguchiID + "' ";
                                    }

                                    Console.WriteLine(cmd.CommandText);
                                    cmd.ExecuteNonQuery();

                                    if (result > 0)
                                    {
                                        Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "調査品目の項目を更新しました。  窓口ID = " + MadoguchiID + " 部所: " + data[i, 1] + " 調査員:" + data[i, 2], "MadoguchiUpdate_SQL", MadoguchiID);
                                    }

                                    DataTable shibuDT = new DataTable();
                                    shibuDT = getData("ShibuBikouChousaBusho", "ShibuBikouChousaBusho", "ShibuBikou with(nolock) ", "ShibuBikouBushoKanriboBushoCD = '" + data[i, 1] + "' AND MadoguchiID = '" + MadoguchiID + "'");
                                    string chousaBusho = "";

                                    if(shibuDT != null && shibuDT.Rows.Count > 0)
                                    {
                                        chousaBusho = shibuDT.Rows[0][0].ToString();

                                        // ShibuBikouChousaBushoに変更があるなら更新する
                                        if (chousaBusho != bikou)
                                        {
                                            cmd.CommandText = "UPDATE ShibuBikou SET " +
                                                        "ShibuBikouChousaBusho = '" + bikou + "' " +
                                                        ",ShibuBikouUpdateDate = SYSDATETIME()  " +
                                                        ",ShibuBikouUpdateUser = N'" + UserInfos[0] + "' " +
                                                        ",ShibuBikouUpdateProgram = '" + pgmName + methodName + "' " +
                                                        " WHERE ShibuBikouBushoKanriboBushoCD = '" + data[i, 1] + "' AND MadoguchiID = '" + MadoguchiID + "' ";

                                            Console.WriteLine(cmd.CommandText);
                                            result = cmd.ExecuteNonQuery();
                                            shibuMessageFlg = 1;
                                        }
                                    }
                                    else
                                    {
                                        cmd.CommandText = "INSERT INTO ShibuBikou ( " +
                                                        "MadoguchiID " +
                                                        ",ShibuBikouID " +
                                                        ",ShibuBikouBushoKanriboBushoCD " +
                                                        ",ShibuBikouKanriNo " +
                                                        ",ShibuBikouChousaBusho " +
                                                        ",ShibuBikou " +
                                                        ",ShibuBikouRyakumei " +
                                                        ",ShibuBikouCreateDate " +
                                                        ",ShibuBikouCreateUser " +
                                                        ",ShibuBikouCreateProgram " +
                                                        ",ShibuBikouUpdateDate " +
                                                        ",ShibuBikouUpdateUser " +
                                                        ",ShibuBikouUpdateProgram " +
                                                        ",ShinDeleteFlag " +
                                                        ",ShibuBushokanriboKameiRakuOld " +
                                                        ",ShibuBushoKanriboShibuMeiOld " +
                                                        ") VALUES  ( " +
                                                        MadoguchiID +
                                                        "," + getSaiban("ShibuBikouID") +
                                                        ", " + data[i, 1] +
                                                        ",NULL " +
                                                        ",'" + bikou + "' " +
                                                        ",NULL " +
                                                        //", " + data[i, 2] +
                                                        ", (SELECT BushokanriboKameiRaku FROM MST_Busho WHERE GyoumuBushoCD = '" + data[i, 1] + "')" +
                                                        ",SYSDATETIME() " +                             // 登録日時
                                                        ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                                                        ",'" + pgmName + methodName + "' " +            // 登録プログラム
                                                        ",SYSDATETIME() " +                             // 更新日時
                                                        ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                                                        ",'" + pgmName + methodName + "' " +            // 更新プログラム
                                                        ",0 " +                                         // 削除フラグ
                                                        ",NULL " +
                                                        ",NULL " +
                                                        " ) ";

                                        Console.WriteLine(cmd.CommandText);
                                        result = cmd.ExecuteNonQuery();
                                        shibuMessageFlg = 1;
                                    }

                                    //cmd.CommandText = "UPDATE ShibuBikou SET " +
                                    //                    "ShibuBikouChousaBusho = '" + bikou + "' " +
                                    //                    ",ShibuBikouUpdateDate = SYSDATETIME()  " +
                                    //                    ",ShibuBikouUpdateUser = '" + UserInfos[0] + "' " +
                                    //                    ",ShibuBikouUpdateProgram = '" + pgmName + methodName + "' " +
                                    //                    " WHERE ShibuBikouBushoKanriboBushoCD = '" + data[i, 1] + "' AND MadoguchiID = '" + MadoguchiID + "' ";

                                    //Console.WriteLine(cmd.CommandText);
                                    //result = cmd.ExecuteNonQuery();

                                    //if (result <= 0)
                                    //{
                                    //    cmd.CommandText = "INSERT INTO ShibuBikou ( " +
                                    //                        "MadoguchiID " +
                                    //                        ",ShibuBikouID " +
                                    //                        ",ShibuBikouBushoKanriboBushoCD " +
                                    //                        ",ShibuBikouKanriNo " +
                                    //                        ",ShibuBikouChousaBusho " +
                                    //                        ",ShibuBikou " +
                                    //                        ",ShibuBikouRyakumei " +
                                    //                        ",ShibuBikouCreateDate " +
                                    //                        ",ShibuBikouCreateUser " +
                                    //                        ",ShibuBikouCreateProgram " +
                                    //                        ",ShibuBikouUpdateDate " +
                                    //                        ",ShibuBikouUpdateUser " +
                                    //                        ",ShibuBikouUpdateProgram " +
                                    //                        ",ShinDeleteFlag " +
                                    //                        ",ShibuBushokanriboKameiRakuOld " +
                                    //                        ",ShibuBushoKanriboShibuMeiOld " +
                                    //                        ") VALUES  ( " +
                                    //                        MadoguchiID +
                                    //                        "," + getSaiban("ShibuBikouID") +
                                    //                        ", " + data[i, 1] +
                                    //                        ",NULL " +
                                    //                        ",'" + bikou + "' " +
                                    //                        ",NULL " +
                                    //                        ", " + data[i, 2] +
                                    //                        ",SYSDATETIME() " +                             // 登録日時
                                    //                        ",'" + UserInfos[0] + "' " +                    // 登録ユーザ
                                    //                        ",'" + pgmName + methodName + "' " +            // 登録プログラム
                                    //                        ",SYSDATETIME() " +                             // 更新日時
                                    //                        ",'" + UserInfos[0] + "' " +                    // 更新ユーザ
                                    //                        ",'" + pgmName + methodName + "' " +            // 更新プログラム
                                    //                        ",0 " +                                         // 削除フラグ
                                    //                        ",NULL " +
                                    //                        ",NULL " +
                                    //                        " ) ";

                                    //    Console.WriteLine(cmd.CommandText);
                                    //    result = cmd.ExecuteNonQuery();
                                    //}


                                    //調査品目明細更新フラグ
                                    messageFlg4 = 1;
                                    //締切日更新フラグ
                                    messageFlg5 = 1;

                                }
                            }
                            //else if (!dt.Rows[0][1].ToString().Equals(data[i, 5]))
                            //{

                            //}
                            //else
                            //{
                            //    mes += GetMessage("I20205", "");
                            //    mes += Environment.NewLine + GetMessage("I20209", "");
                            //}
                        }
                        else
                        {
                            //調査担当者データが存在しない
                            transaction.Rollback();
                            sqlconn.Close();
                            return false;
                        }
                    }

                    //進捗状況の最小値を取得
                    //cmd.CommandText = "SELECT MAX(CASE MadoguchiL1ChousaShinchoku WHEN 7 THEN 2.5 ELSE MadoguchiL1ChousaShinchoku END) AS 'MAX' " +　//進捗状況のデータ2桁変更
                    //                    ",MIN(CASE MadoguchiL1ChousaShinchoku WHEN 7 THEN 2.5 ELSE MadoguchiL1ChousaShinchoku END) AS 'MIN' " +
                    cmd.CommandText = "SELECT MAX(MadoguchiL1ChousaShinchoku ) AS 'MAX' " +
                                        ",MIN(MadoguchiL1ChousaShinchoku ) AS 'MIN' " +

                                        " FROM MadoguchiJouhouMadoguchiL1Chou WHERE MadoguchiID = " + MadoguchiID;

                    var dt2 = new DataTable();
                    var sda2 = new SqlDataAdapter(cmd);
                    sda2.Fill(dt2);

                    if (dt2 != null && dt2.Rows.Count > 0 && dt2.Rows[0][1] != null && dt2.Rows[0][1].ToString() != "")
                    {
                        //進捗状況のデータ2桁変更により、進捗状況の順番が正常化する
                        //if (dt2.Rows[0][1].Equals("2.5"))
                        //{
                        //    cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                        //                        //"MadoguchiShinchokuJoukyou = '7' " +
                        //                        "MadoguchiShinchokuJoukyou = '60' " +
                        //                        " WHERE MadoguchiID = " + MadoguchiID;
                        //}
                        //else
                        //{
                        cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                                            "MadoguchiShinchokuJoukyou = '" + (int)double.Parse(dt2.Rows[0][1].ToString()) + "' " +
                                            ",MadoguchiUpdateDate = SYSDATETIME() " +
                                            ",MadoguchiUpdateUser = N'" + UserInfos[0] + "' " +
                                            ",MadoguchiUpdateProgram = '" + pgmName + methodName + "' " +
                                            " WHERE MadoguchiID = " + MadoguchiID;
                        //}
                        Console.WriteLine(cmd.CommandText);
                        var result = cmd.ExecuteNonQuery();

                        outputLogger("MadoguchiUpdate_SQL", "窓口ID = " + MadoguchiID + " 最小値 = " + dt2.Rows[0][1] + " 最大値 = " + dt2.Rows[0][0], "進捗状況更新", "DEBUG");
                    }


                    //フラグ別メッセージ取得
                    //messageFlg1,messageFlg2,messageFlg3は削除時のメッセージの為スルー
                    //messageFlg3(削除処理で削除データがないとき) 
                    //if (messageFlg == 1)
                    //{
                    //    //調査員更新
                    //    // I20206:調査員が更新されました。
                    //    mes += GetMessage("I20206", "");

                    //}

                    //if (messageFlg4 == 1)
                    //{
                    //    //調査品目更新
                    //    // I20202:調査品目明細を更新しました。
                    //    mes += Environment.NewLine + GetMessage("I20202", "");

                    //}

                    //if (messageFlg5 == 1)
                    //{
                    //    //締切日更新
                    //    // I20208:締切日が変更されました。更新を行うと変更した担当者の締切日が全て変更となります。
                    //    mes += Environment.NewLine + GetMessage("I20208", "");
                    //}

                    // 1188 メッセージは1つにする
                    if (messageFlg == 1)
                    {
                        // I30206:担当者状況を変更しました。
                        mes += Environment.NewLine + GetMessage("I30206", "");
                    }
                    if (shibuMessageFlg == 1)
                    {
                        // I20207:支部備考を更新しました。
                        mes += Environment.NewLine + GetMessage("I20207", "");
                    }

                    ////すべてのフラグが0のとき
                    //if (messageFlg == 0 &&
                    //    messageFlg1 == 0 &&
                    //    messageFlg2 == 0 &&
                    //    messageFlg3 == 0 &&
                    //    messageFlg4 == 0 &&
                    //    messageFlg5 == 0)
                    //{
                    //    //更新された情報はありません。
                    //    mes += GetMessage("I20209", "");
                    //}

                    // Garoon追加宛先追加メッセージフラグ
                    int garoonMessageFlg1 = 0;
                    // Garoon追加宛先更新メッセージフラグ
                    int garoonMessageFlg2 = 0;
                    // Garoon追加宛先削除メッセージフラグ
                    int garoonMessageFlg3 = 0;

                    // Garoon追加宛先のテーブルを消すのはなし
                    //Garoon追加宛先更新
                    //cmd.CommandText = "DELETE FROM GaroonTsuikaAtesaki " +
                    //                    " WHERE GaroonTsuikaAtesakiMadoguchiID = " + MadoguchiID;

                    //Console.WriteLine(cmd.CommandText);
                    //cmd.ExecuteNonQuery();


                    var GaroonTsuikaAtesakiBushoCDList = new List<string>();
                    var GaroonTsuikaAtesakiBushoList = new List<string>();
                    var GaroonTsuikaAtesakiTantoushaCDList = new List<string>();
                    var GaroonTsuikaAtesakiTantoushaList = new List<string>();
                    var GaroonTsuikaAtesakiIDList = new List<string>();

                    var i_keyList_c = new List<string>();
                    var i_modeList_c = new List<string>();

                    // 現在のGaroon追加宛先テーブルの情報を取得
                    cmd.CommandText = "SELECT " +
                                      " GaroonTsuikaAtesakiBushoCD " +
                                      ",GaroonTsuikaAtesakiBusho " +
                                      ",GaroonTsuikaAtesakiTantoushaCD " +
                                      ",GaroonTsuikaAtesakiTantousha " +
                                      ",GaroonTsuikaAtesakiID " +
                                      "FROM GaroonTsuikaAtesaki " +
                                      " WHERE GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "'  AND GaroonTsuikaAtesakiDeleteFlag <> 1 ";

                    var dtGaroonTsuikaAtesaki = new DataTable();
                    var sdaGaroonTsuikaAtesaki = new SqlDataAdapter(cmd);
                    sdaGaroonTsuikaAtesaki.Fill(dtGaroonTsuikaAtesaki);

                    // 取得データがある場合
                    if (dtGaroonTsuikaAtesaki.Rows.Count > 0)
                    {
                        for (int i = 0; dtGaroonTsuikaAtesaki.Rows.Count > i; i++)
                        {
                            // Listに詰める
                            GaroonTsuikaAtesakiBushoCDList.Add(dtGaroonTsuikaAtesaki.Rows[i][0].ToString());
                            GaroonTsuikaAtesakiBushoList.Add(dtGaroonTsuikaAtesaki.Rows[i][1].ToString());
                            GaroonTsuikaAtesakiTantoushaCDList.Add(dtGaroonTsuikaAtesaki.Rows[i][2].ToString());
                            GaroonTsuikaAtesakiTantoushaList.Add(dtGaroonTsuikaAtesaki.Rows[i][3].ToString());
                            GaroonTsuikaAtesakiIDList.Add(dtGaroonTsuikaAtesaki.Rows[i][4].ToString());

                            // Keyを設定
                            i_keyList_c.Add(dtGaroonTsuikaAtesaki.Rows[i][0].ToString() + "," + dtGaroonTsuikaAtesaki.Rows[i][2].ToString()); // GaroonTsuikaAtesakiBushoCD + "," + GaroonTsuikaAtesakiTantoushaCD
                            // 2:更新 3:削除 それ以外は新規
                            i_modeList_c.Add("3");
                        }
                    }

                    // 登録フラグ true:登録する false:登録なし
                    Boolean insertFlg = true;
                    string i_key = "";
                    for (int i = 0; i < data2.GetLength(0); i++)
                    {
                        // data2
                        // 0:GaroonTsuikaAtesakiID
                        // 1:GaroonTsuikaAtesakiBushoCD
                        // 2:GaroonTsuikaAtesakiBusho
                        // 3:GaroonTsuikaAtesakiTantoushaCD
                        // 4:GaroonTsuikaAtesakiTantousha

                        insertFlg = true;

                        i_key = data2[i, 1] + "," + data2[i, 3];

                        // テーブルの中身と一致した場合、既に存在する為、更新なし
                        for (int j = 0; j < i_keyList_c.Count; j++)
                        {
                            // 部所CD、名前が一致するか
                            if (i_keyList_c[j] == i_key)
                            {
                                insertFlg = false;

                                // 差異があった場合に更新する
                                //if (data2[i, 2] != GaroonTsuikaAtesakiBushoList[i] || data2[i, 4] != GaroonTsuikaAtesakiTantoushaList[i]) { 
                                // ProUpdateTantoubusho と新で入れてる名称が違うので、CDで比較する
                                if (data2[i, 1] != GaroonTsuikaAtesakiBushoCDList[i] || data2[i, 3] != GaroonTsuikaAtesakiTantoushaCDList[i]) { 
                                    cmd.CommandText = "UPDATE GaroonTsuikaAtesaki SET " +
                                                      " GaroonTsuikaAtesakiBushoCD = '" + data2[i, 1] + "' " +
                                                      ",GaroonTsuikaAtesakiBusho = N'" + data2[i, 2] + "' " +
                                                      ",GaroonTsuikaAtesakiTantoushaCD = '" + data2[i, 3] + "' " +
                                                      ",GaroonTsuikaAtesakiTantousha = N'" + data2[i, 4] + "' " +
                                                      ",GaroonTsuikaAtesakiUpdateDate = SYSDATETIME() " +
                                                      ",GaroonTsuikaAtesakiUpdateUser = N'" + UserInfos[0] + "' " +
                                                      ",GaroonTsuikaAtesakiUpdateProgram = '" + pgmName + methodName + "'" +
                                                      //不具合No1332(1084) 画面登録か品目追加により書き換える
                                                      ",GaroonTsuikaAtesakiGamenFlag = '" + data2[i, 5] + "'" +
                                                      " WHERE GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "' " +
                                                      "AND GaroonTsuikaAtesakiBushoCD = '" + data2[i, 1] + "' " +
                                                      "AND GaroonTsuikaAtesakiTantoushaCD = '" + data2[i, 3] + "' ";
                                    cmd.ExecuteNonQuery();

                                    // Garoon追加宛先更新メッセージフラグ
                                    // 部所、担当者がの名前が変わった場合しか更新されない
                                    garoonMessageFlg2 = 1;
                                }

                                //// 既に存在している（削除もなし）なので、Listから除外
                                //GaroonTsuikaAtesakiBushoCDList.RemoveAt(j);
                                //GaroonTsuikaAtesakiBushoList.RemoveAt(j);
                                //GaroonTsuikaAtesakiTantoushaCDList.RemoveAt(j);
                                //GaroonTsuikaAtesakiTantoushaList.RemoveAt(j);
                                //GaroonTsuikaAtesakiIDList.RemoveAt(j);
                                // 更新
                                i_modeList_c[j] = "2";
                            }
                        }

                        string GaroonTsuikaAtesakiID = "";
                        if (insertFlg == true)
                        {
                            //担当者が未選択の場合は登録なし
                            if (data2[i, 3] != null && data2[i, 3] != "")
                            {

                                GaroonTsuikaAtesakiID = GaroonTsuikaAtesakiID = getSaiban("GaroonTsuikaAtesakiID").ToString();

                                cmd.CommandText = "INSERT INTO GaroonTsuikaAtesaki ( " +
                                                " GaroonTsuikaAtesakiID " +
                                                ",GaroonTsuikaAtesakiMadoguchiID " +
                                                ",GaroonTsuikaAtesakiBushoCD " +
                                                ",GaroonTsuikaAtesakiBusho " +
                                                ",GaroonTsuikaAtesakiTantoushaCD " +
                                                ",GaroonTsuikaAtesakiTantousha " +
                                                ",GaroonTsuikaAtesakiCreateDate " +
                                                ",GaroonTsuikaAtesakiCreateUser " +
                                                ",GaroonTsuikaAtesakiCreateProgram " +
                                                ",GaroonTsuikaAtesakiUpdateDate " +
                                                ",GaroonTsuikaAtesakiUpdateUser " +
                                                ",GaroonTsuikaAtesakiUpdateProgram " +
                                                ",GaroonTsuikaAtesakiDeleteFlag " +
                                                //不具合No1332(1084) 
                                                ",GaroonTsuikaAtesakiGamenFlag " +
                                                " ) VALUES ( " +
                                                " '" + GaroonTsuikaAtesakiID + "' " +
                                                ",'" + MadoguchiID + "' " +
                                                ",'" + data2[i, 1] + "' " +
                                                ",N'" + data2[i, 2] + "' " +
                                                ",'" + data2[i, 3] + "' " +
                                                ",N'" + data2[i, 4] + "' " +
                                                ",SYSDATETIME() " +                             // 登録日時
                                                ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                                                ",'" + pgmName + methodName + "' " +            // 登録プログラム
                                                ",SYSDATETIME() " +                             // 更新日時
                                                ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                                                ",'" + pgmName + methodName + "' " +            // 更新プログラム
                                                ",0 " +                                         // 削除フラグ
                                                //不具合No1332(1084) 
                                                ",'" + data2[i, 5] + "' " +                     //画面から追加されたかフラグ
                                                " ) ";
                                Console.WriteLine(cmd.CommandText);
                                cmd.ExecuteNonQuery();

                                // Garoon追加宛先追加メッセージフラグ
                                garoonMessageFlg1 = 1;
                            }
                        }
                    }
                    // 画面から削除されたデータをテーブル上からも削除する
                    for (int i = 0; i < i_modeList_c.Count; i++)
                    {
                        if ("3".Equals(i_modeList_c[i]))
                        {
                            cmd.CommandText = "DELETE FROM GaroonTsuikaAtesaki " +
                                              " WHERE GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "' " +
                                              "AND GaroonTsuikaAtesakiID = '" + GaroonTsuikaAtesakiIDList[i] + "' ";

                            cmd.ExecuteNonQuery();
                            // Garoon追加宛先削除メッセージフラグ
                            garoonMessageFlg3 = 1;
                        }
                    }

                    //if (garoonMessageFlg1 == 1)
                    //{
                    //    // Garoon追加宛先追加
                    //    // I30203:Garoon追加宛先を追加しました。
                    //    mes += Environment.NewLine + GetMessage("I30203", "");
                    //}

                    //if (garoonMessageFlg2 == 1)
                    //{
                    //    // Garoon追加宛先更新
                    //    // I30204:Garoon追加宛先を更新しました。
                    //    mes += Environment.NewLine + GetMessage("I30204", "");
                    //}

                    //if (garoonMessageFlg3 == 1)
                    //{
                    //    // Garoon追加宛先削除
                    //    // I30205:Garoon追加宛先を削除しました。
                    //    mes += Environment.NewLine + GetMessage("I30205", "");
                    //}
                    // 1188 メッセージは1つ
                    if (garoonMessageFlg1 == 1 || garoonMessageFlg2 == 1 || garoonMessageFlg3 == 1)
                    {
                        // I30204:Garoon追加宛先を更新しました。
                        mes += Environment.NewLine + GetMessage("I30204", "");
                    }

                    //すべてのフラグが0のとき
                    if (messageFlg == 0 &&
                        messageFlg1 == 0 &&
                        messageFlg2 == 0 &&
                        messageFlg3 == 0 &&
                        messageFlg4 == 0 &&
                        messageFlg5 == 0 &&
                        garoonMessageFlg1 == 0 &&
                        garoonMessageFlg2 == 0 &&
                        garoonMessageFlg3 == 0 &&
                        shibuMessageFlg == 0)
                    {
                        //更新された情報はありません。
                        mes += GetMessage("I20209", "");
                    }

                    //// 皇帝まもる連携
                    //KouteiTantouBushoRenkei(MadoguchiID, UserInfos[0], UserInfos[2]);

                }
                // 協力依頼書
                else if (tab == 4)
                {
                    cmd.CommandText = "UPDATE KyouryokuIraisho SET ";
                    if (data[0, 0] != null && data[0, 0] != "" && data[0, 0] != "null")
                    {
                        cmd.CommandText += " KyourokuIraisakiBushoOld      = '" + data[0, 0] + "' ";
                    }
                    else
                    {
                        cmd.CommandText += " KyourokuIraisakiBushoOld      = NULL ";
                    }
                    if (data[0, 23] != null && data[0, 23] != "")
                    {
                        cmd.CommandText += ",KyourokuIraisakiTantoshaCD    = '" + data[0, 23] + "' ";
                    }
                    else
                    {
                        cmd.CommandText += ",KyourokuIraisakiTantoshaCD                 = NULL ";
                    }
                    if (data[0, 1] != null && data[0, 1] != "")
                    {
                        cmd.CommandText += ",KyouryokuDate                 = '" + data[0, 1] + "' ";
                    }
                    else
                    {
                        cmd.CommandText += ",KyouryokuDate                 = NULL ";
                    }
                    if (data[0, 2] != null && data[0, 2] != "")
                    {
                        cmd.CommandText += ",KyouryokuHoukokuSeigenDate     = '" + data[0, 2] + "' ";
                    }
                    else
                    {
                        cmd.CommandText += ",KyouryokuHoukokuSeigenDate     = NULL ";
                    }

                    //",KyouryokuOuenirai             = '" + data[0, 0] + "' " + //画面項目削除
                    cmd.CommandText += ",KyouryokuGyoumuKubun          = N'" + data[0, 3] + "' " +
                                            ",KyouryokuIraiKubun            = N'" + data[0, 4] + "' " +
                                            ",KyouryokuNaiyoukubunShizai    = N'" + data[0, 5] + "' " +
                                            ",KyouryokuNaiyoukubunDKou      = N'" + data[0, 6] + "' " +
                                            ",KyouryokuNaiyoukubunEKou      = N'" + data[0, 7] + "' " +
                                            ",KyouryokuNaiyoukubunSonota    = N'" + data[0, 8] + "' " +
                                            ",KyouryokuNaiyoukubunJohokaihatsu  = N'" + data[0, 9] + "' " +
                                            ",KyouryokuSonota               = N'" + ChangeSqlText(data[0, 10], 0) + "' " +
                                            ",KyouryokuGyoumuNaiyou         = N'" + ChangeSqlText(data[0, 11], 0) + "' " +
                                            ",KyouryokuZumen                = N'" + data[0, 12] + "' " +
                                            ",KyouryokuChousaKijun          = N'" + data[0, 13] + "' " +
                                            ",KyouryokuChousakijunbi        = N'" + ChangeSqlText(data[0, 14], 0) + "' " +
                                            ",KyouryokuUtiawaseyouhi        = N'" + data[0, 15] + "' " +
                                            ",KyouryokuGutaiteki            = N'" + data[0, 16] + "' " +
                                            ",KyouryokuZenkaiUmu            = N'" + data[0, 17] + "' " +
                                            ",KyouryokuZenkaiUmubi          = N'" + ChangeSqlText(data[0, 18], 0) + "' " +
                                            ",KyouryokusakiHikiwatashi      = N'" + data[0, 19] + "' " +
                                            ",KyouryokuJisshikeikakusho     = N'" + data[0, 20] + "' " +
                                            ",KyouryokuChoushuusaki         = N'" + data[0, 21] + "' " +
                                            //",KyouryokuRenrakuJikou         = '' " +
                                            //",KyouryokuIraishoHozonsaki     = '" + data[0, 0] + "' " +
                                            ",KyouryokuUpdateDate           = SYSDATETIME() " +
                                            ",KyouryokuUpdateUser           = N'" + UserInfos[0] + "' " +
                                            ",KyouryokuUpdateProgram        = '" + pgmName + methodName + "' " +
                                            ",KyouryokuDeleteFlag           = '0' " +
                                        " WHERE KyouryokuIraishoID = '" + data[0, 22] + "' AND MadoguchiID = '" + MadoguchiID + "' ";
                    Console.WriteLine(cmd.CommandText);
                    var result = cmd.ExecuteNonQuery();

                    if (result > 0)
                    {
                        Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "協力依頼書画面を更新しました。 窓口ID = " + MadoguchiID, "MadoguchiUpdate_SQL", MadoguchiID);

                        // I20404：データを更新しました。
                        mes += GetMessage("I20404", "");

                    }
                    else
                    {
                        // E20402:更新に失敗しました。
                        mes += GetMessage("E20402", "");
                    }

                }
                // 応援受付
                else if (tab == 5)
                {
                    cmd.CommandText = "UPDATE OuenUketsuke SET ";
                    // 受付状況
                    if (data[0, 0] != null && data[0, 0] == "1")
                    {
                        cmd.CommandText += " OuenJoukyou = 1 ";
                    }
                    else if (data[0, 0] != null && data[0, 0] == "2")
                    {
                        cmd.CommandText += " OuenJoukyou = 2 ";
                    }
                    else
                    {
                        cmd.CommandText += " OuenJoukyou = 0 ";
                    }
                    // 応援受付日
                    if (data[0, 1] != null && data[0, 1] != "")
                    {
                        cmd.CommandText += ",OuenUketsukeDate = '" + data[0, 1] + "' ";
                    }
                    else
                    {
                        cmd.CommandText += ",OuenUketsukeDate = NULL ";
                    }
                    // 応援完了
                    if (data[0, 2] != null && data[0, 2] == "1")
                    {
                        cmd.CommandText += ",OuenKanryou = 1 ";
                    }
                    else
                    {
                        cmd.CommandText += ",OuenKanryou = 0 ";
                    }
                    // 応援完了日
                    if (data[0, 3] != null && data[0, 3] != "")
                    {
                        cmd.CommandText += ",OuenHoukokuJishibi = '" + data[0, 3] + "' ";
                    }
                    else
                    {
                        cmd.CommandText += ",OuenHoukokuJishibi = NULL ";
                    }
                    cmd.CommandText += ",OuenUpdateDate    = SYSDATETIME() " +
                                       ",OuenUpdateUser    = N'" + UserInfos[0] + "' " +
                                       ",OuenUpdateProgram = '" + pgmName + methodName + "' " +
                                       ",OuenDeleteFlag    = 0 " +
                                       ",Ouen10 = 1 " +
                                       " WHERE MadoguchiID = '" + MadoguchiID + "' AND OuenDeleteFlag != 1 ";
                    //Console.WriteLine(cmd.CommandText);
                    var result = cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE MadoguchiJouhou SET ";
                    // 応援受付日
                    if (data[0, 1] != null && data[0, 1] != "")
                    {
                        cmd.CommandText += "MadoguchiOuenUketsukebi = '" + data[0, 1] + "' ";
                    }
                    else
                    {
                        cmd.CommandText += "MadoguchiOuenUketsukebi = NULL ";
                    }
                    cmd.CommandText += ",MadoguchiUpdateDate    = SYSDATETIME() " +
                                       ",MadoguchiUpdateUser    = N'" + UserInfos[0] + "' " +
                                       ",MadoguchiUpdateProgram = '" + pgmName + methodName + "' " +
                                       " WHERE MadoguchiID = '" + MadoguchiID + "' AND MadoguchiDeleteFlag != 1 ";
                    var result2 = cmd.ExecuteNonQuery();

                    transaction.Commit();

                    if (result > 0)
                    {
                        Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "応援受付状況を更新しました。 窓口ID = " + MadoguchiID, "MadoguchiUpdate_SQL", MadoguchiID);

                        // データを更新しました。
                        mes += GetMessage("I20404", "");

                    }
                    else
                    {
                        // 更新に失敗しました。
                        mes += GetMessage("E20402", "");
                    }
                }
                // 単品入力情報
                else if (tab == 6)
                {
                    String strSpace = " ";
                    String strComma = ", ";
                    String strEqual = " = ";
                    string strSingleQuote = "'";
                    int result = 0;

                    StringBuilder sb = new StringBuilder();

                    // テーブル定義（項目名、属性）
                    string[,] TanpinNyuuryoku = new string[,]
                    {
                         {"TanpinNyuuryokuID", "Numeric"}	            // 0.単品入力項目ID
                        ,{"TanpinJutakuDate", "Date"}	                // 1.受託日（依頼日）
                        ,{"TanpinHoukokuDate", "Date"}	                // 2.報告日
                        ,{"TanpinShiji", "String"}	                    // 3.指示番号
                        ,{"TanpinHachuubusho", "String"}	            // 4.部所
                        ,{"TanpinYakushoku", "String"}	                // 5.役職
                        ,{"TanpinHachuuTantousha", "String"}	        // 6.担当者
                        ,{"TanpinTel", "String"}	                    // 7.電話
                        ,{"TanpinFax", "String"}	                    // 8.FAX
                        ,{"TanpinMail", "String"}	                    // 9.メール
                        ,{"TanpinMemo", "String"}	                    // 10.メモ
                        ,{"TanpinRank", "String"}	                    // 11.ランク
                        ,{"TanpinShousa", "String"}	                    // 12.照査実施
                        ,{"TanpinShijisho", "Numeric"}	                // 13.指示書
                        ,{"TanpinSaishuuKensa", "Numeric"}	            // 14.最終検査
                        ,{"TanpinMitsumoriTeishutu", "Numeric"}	        // 15.見積提出方式
                        ,{"TanpinTeinyuusatsu", "Numeric"}	            // 16.低入札
                        ,{"TanpinShuyouChousain", "String"}	            // 17.主要調査員
                        ,{"TanpinSeikyuuGetsu", "String"}	            // 18.単品請求月
                        ,{"TanpinHokurikuShijouKakaku", "Numeric"}	    // 19.市場価格（北陸専用）
                        ,{"TanpinHokurikuShijouKakaku_r", "Numeric"}	// 20.市場価格（北陸専用）r
                        ,{"TanpinHokurikuSekouKanka", "Numeric"}	    // 21.施工単価（北陸専用）
                        ,{"TanpinHokurikuSekouKanka_r", "Numeric"}	    // 22.施工単価（北陸専用）r
                        ,{"TanpinSonotaShuukei", "Numeric"}	            // 23.その他集計
                        ,{"TanpinSeikyuuKingaku", "Numeric"}	        // 24.請求金額
                        ,{"TanpinSeikyuuKakutei", "Numeric"}	        // 25.請求確定
                        ,{"MadoguchiID", "Numeric"}	                    // 26.窓口ID
                        ,{"TanpinGyoumuCD", "Numeric"}	                // 27.業務CD
                        ,{"TanpinAnkenJouhouID", "Numeric"}	            // 28.契約情報ID
                        ,{"TanpinKeihi", "Numeric"}	                    // 29.経費（バックアップ用）
                        ,{"TanpinCreateDate", "Date"}	                // 30.作成日時
                        ,{"TanpinCreateUser", "String"}	                // 31.作成ユーザ
                        ,{"TanpinCreateProgram", "String"}	            // 32.作成機能
                        ,{"TanpinUpdateDate", "Date"}	                // 33.更新日時
                        ,{"TanpinUpdateUser", "String"}	                // 34.更新ユーザ
                        ,{"TanpinUpdateProgram", "String"}	            // 35.更新機能
                        ,{"TanpinDeleteFlag", "Numeric"}                // 36.削除フラグ
                    };

                    string[,] TanpinNyuuryokuRank = new string[,]
                    {
                         {"TanpinNyuuryokuID", "Numeric"}               // 0.単品入力項目ID
                        ,{"TanpinL1RankID", "Numeric"}                  // 1.ランクID
                        ,{"TanpinL1RankMei", "String"}                  // 2.ランク名
                        ,{"TanpunL1RankKubun", "Numeric"}               // 3.ランク種別（集計方法）
                        ,{"TanpinL1Ranksuu", "Numeric"}                 // 4.依頼本数
                        ,{"TanpinL1HoukokuHonsuu", "Numeric"}           // 5.報告本数
                        ,{"TanpinL1Tanka", "Numeric"}                   // 6.単価
                        ,{"TanpinL1Kingaku", "Numeric"}                 // 7.金額
                    };

                    // 対象データの存在チェック
                    var dt = new DataTable();
                    string TanpinNyuuryokuID = data[0, 0];
                    // SQL生成
                    cmd.CommandText = "SELECT TanpinNyuuryokuID FROM TanpinNyuuryoku WHERE TanpinNyuuryokuID = " + TanpinNyuuryokuID;
                    // データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);

                    // 受託番号で単価契約の紐づきが変わると前のデータが残るのでランクはDelete、Insertにする
                    cmd.CommandText = "DELETE FROM TanpinNyuuryokuRank WHERE TanpinNyuuryokuID = " + TanpinNyuuryokuID;
                    cmd.ExecuteNonQuery();

                    // データが存在したら更新、なければ登録
                    if (dt.Rows.Count > 0)
                    {
                        // UPDATE文の生成
                        sb.Clear();
                        sb.Append("UPDATE TanpinNyuuryoku SET ");
                        for (int i = 0; i < TanpinNyuuryoku.GetLength(0); i++)
                        {
                            if (i == 0 || i == 1 || i == 30 || i == 31 || i == 32)  // UPDATE対象外項目
                            {
                            }
                            else
                            {
                                sb.Append(strComma);
                            }
                            switch (i)
                            {
                                case 0: // 単品入力項目ID
                                    break;
                                case 30: // 作成日時
                                case 31: // 作成ユーザ
                                case 32: // 作成機能
                                    break;
                                case 33: // 更新日時
                                    sb.Append(TanpinNyuuryoku[i, 0]);
                                    sb.Append(strEqual);
                                    sb.Append("SYSDATETIME()");
                                    break;
                                case 34: // 更新ユーザ
                                    sb.Append(TanpinNyuuryoku[i, 0]);
                                    sb.Append(strEqual);
                                    sb.Append(strSingleQuote);
                                    sb.Append(UserInfos[0]);
                                    sb.Append(strSingleQuote);
                                    break;
                                case 35: // 更新機能
                                    sb.Append(TanpinNyuuryoku[i, 0]);
                                    sb.Append(strEqual);
                                    sb.Append(strSingleQuote);
                                    sb.Append(pgmName + methodName);
                                    sb.Append(strSingleQuote);
                                    break;
                                default:
                                    sb.Append(TanpinNyuuryoku[i, 0]);
                                    sb.Append(strEqual);
                                    // データの属性によって処理を分ける
                                    switch (TanpinNyuuryoku[i, 1])
                                    {
                                        case "String":
                                            sb.Append("N" + strSingleQuote);
                                            sb.Append(data[0, i]);
                                            sb.Append(strSingleQuote);
                                            break;
                                        case "Numeric":
                                            if (data[0, i] != null && data[0, i] != "")
                                            {
                                                sb.Append(GetLong(data[0, i]));
                                            }
                                            else
                                            {
                                                sb.Append("0");
                                            }
                                            break;
                                        case "Date":
                                            if (data[0, i] != null && data[0, i] != "")
                                            {
                                                sb.Append(strSingleQuote);
                                                sb.Append(data[0, i]);
                                                sb.Append(strSingleQuote);
                                            }
                                            else
                                            {
                                                sb.Append("NULL");
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                    break;
                            }
                        }

                        // 条件式の設定
                        sb.Append(" WHERE ");
                        sb.Append(TanpinNyuuryoku[0, 0]);   // 単品入力項目ID
                        sb.Append(strEqual);
                        sb.Append(TanpinNyuuryokuID);

                        cmd.CommandText = sb.ToString();

                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();

                        // I20602:単品入力項目が新規で作成されました。
                        //mes = GetMessage("I20602", "");

                    }
                    else
                    {
                        // INSERT文の生成
                        sb.Clear();
                        sb.Append("INSERT INTO TanpinNyuuryoku ( ");
                        // 項目設定
                        for (int i = 0; i < TanpinNyuuryoku.GetLength(0); i++)
                        {
                            if (i != 0)
                            {
                                sb.Append(strComma);
                            }
                            sb.Append(TanpinNyuuryoku[i, 0]);

                        }
                        sb.Append(" ) VALUES ( ");
                        // データセット
                        for (int i = 0; i < TanpinNyuuryoku.GetLength(0); i++)
                        {
                            if (i != 0)
                            {
                                sb.Append(strComma);
                            }
                            switch (i)
                            {
                                case 0: // 単品入力項目ID
                                    TanpinNyuuryokuID = getSaiban("TanpinNyuuryokuID").ToString();
                                    sb.Append(TanpinNyuuryokuID);
                                    break;
                                case 30: // 作成日時
                                case 33: // 更新日時
                                    sb.Append("SYSDATETIME()");
                                    break;
                                case 31: // 作成ユーザ
                                case 34: // 更新ユーザ
                                    sb.Append(strSingleQuote);
                                    sb.Append(UserInfos[0]);
                                    sb.Append(strSingleQuote);
                                    break;
                                case 32: // 作成機能
                                case 35: // 更新機能
                                    sb.Append(strSingleQuote);
                                    sb.Append(pgmName + methodName);
                                    sb.Append(strSingleQuote);
                                    break;
                                case 36: // 削除フラグ
                                    sb.Append("0");
                                    break;
                                default:
                                    // データの属性によって処理を分ける
                                    switch (TanpinNyuuryoku[i, 1])
                                    {
                                        case "String":
                                            sb.Append("N" + strSingleQuote);
                                            sb.Append(data[0, i]);
                                            sb.Append(strSingleQuote);
                                            break;
                                        case "Numeric":
                                            if (data[0, i] != null && data[0, i] != "")
                                            {
                                                sb.Append(GetLong(data[0, i]));
                                            }
                                            else
                                            {
                                                sb.Append("0");
                                            }
                                            break;
                                        case "Date":
                                            if (data[0, i] != null && data[0, i] != "")
                                            {
                                                sb.Append(strSingleQuote);
                                                sb.Append(data[0, i]);
                                                sb.Append(strSingleQuote);
                                            }
                                            else
                                            {
                                                sb.Append("NULL");
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                    break;
                            }

                        }
                        sb.Append(")");
                        cmd.CommandText = sb.ToString();

                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();

                        // I20601:単品入力項目が更新されました。
                        //mes = GetMessage("I20601", "");

                    }

                    // 対象データの存在チェック
                    dt = new DataTable();
                    // SQL生成
                    cmd.CommandText = "SELECT TanpinL1RankID FROM TanpinNyuuryokuRank WHERE TanpinNyuuryokuID = " + TanpinNyuuryokuID;
                    // データ取得
                    sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);

                    string TanpinL1RankID = "";
                    string textInsertSQL = "";
                    for (int i = 0; i < data2.GetLength(0); i++)
                    {
                        // 一括登録用、nullの場合、処理を抜ける
                        if(data2[i, 1] == null)
                        {
                            break;
                        }

                        TanpinL1RankID = data2[i, 1];   // ランクID
                        decimal.TryParse(TanpinL1RankID, out decimal decimalTanpinL1RankID);
                        // 存在したら更新、なければ登録
                        DataRow[] dRows = dt.AsEnumerable().Where(row => row.Field<decimal>("TanpinL1RankID") == decimalTanpinL1RankID).ToArray();

                        if (dRows.Length > 0)
                        {
                            // UPDATE文の生成(都度実行させる)
                            sb.Clear();
                            sb.Append("UPDATE TanpinNyuuryokuRank SET ");
                            for (int j = 0; j < TanpinNyuuryokuRank.GetLength(0); j++)
                            {
                                if (j == 0 || j == 1 || j == 2)
                                {
                                }
                                else
                                {
                                    sb.Append(strComma);
                                }
                                switch (j)
                                {
                                    case 0: // 単品入力項目ID
                                    case 1: // ランクID
                                        break;
                                    default:
                                        sb.Append(TanpinNyuuryokuRank[j, 0]);
                                        sb.Append(strEqual);
                                        // データの属性によって処理を分ける
                                        switch (TanpinNyuuryokuRank[j, 1])
                                        {
                                            case "String":
                                                sb.Append("N" + strSingleQuote);
                                                sb.Append(data2[i, j]);
                                                sb.Append(strSingleQuote);
                                                break;
                                            case "Numeric":
                                                if (data2[i, j] != null && data2[i, j] != "")
                                                {
                                                    sb.Append(GetLong(data2[i, j]));
                                                }
                                                else
                                                {
                                                    sb.Append("0");
                                                }
                                                break;
                                            case "Date":
                                                if (data2[i, j] != null && data2[i, j] != "")
                                                {
                                                    sb.Append(strSingleQuote);
                                                    sb.Append(data2[i, j]);
                                                    sb.Append(strSingleQuote);
                                                }
                                                else
                                                {
                                                    sb.Append("NULL");
                                                }
                                                break;
                                            default:
                                                break;
                                        }
                                        break;
                                }
                            }

                            // 条件式の設定
                            sb.Append(" WHERE ");
                            sb.Append(TanpinNyuuryokuRank[0, 0]);   // 単品入力項目ID
                            sb.Append(strEqual);
                            sb.Append(TanpinNyuuryokuID);
                            sb.Append(strSpace);
                            sb.Append("AND ");
                            sb.Append(TanpinNyuuryokuRank[1, 0]);   // ランクID
                            sb.Append(strEqual);
                            sb.Append(TanpinL1RankID);

                            cmd.CommandText = sb.ToString();

                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();

                        }
                        else
                        {
                            // INSERT文の生成(複数まとめて最後に実行)
                            sb.Clear();

                            if (textInsertSQL == "")
                            {
                                sb.Append("INSERT INTO TanpinNyuuryokuRank ( ");
                                // 項目設定
                                for (int j = 0; j < TanpinNyuuryokuRank.GetLength(0); j++)
                                {
                                    if (j != 0)
                                    {
                                        sb.Append(strComma);
                                    }
                                    sb.Append(TanpinNyuuryokuRank[j, 0]);

                                }
                                sb.Append(" ) VALUES ( ");
                            }
                            else
                            {
                                sb.Append(strComma);
                                sb.Append("( ");
                            }

                            // データセット
                            for (int j = 0; j < TanpinNyuuryokuRank.GetLength(0); j++)
                            {
                                if (j != 0)
                                {
                                    sb.Append(strComma);
                                }
                                switch (j)
                                {
                                    case 0: // 単品入力項目ID
                                        sb.Append(TanpinNyuuryokuID);
                                        break;
                                    case 1: // ランクID
                                        TanpinL1RankID = getSaiban("TanpinL1RankID").ToString();
                                        sb.Append(TanpinL1RankID);
                                        break;
                                    default:
                                        // データの属性によって処理を分ける
                                        switch (TanpinNyuuryokuRank[j, 1])
                                        {
                                            case "String":
                                                sb.Append("N" + strSingleQuote);
                                                sb.Append(data2[i, j]);
                                                sb.Append(strSingleQuote);
                                                break;
                                            case "Numeric":
                                                if (data2[i, j] != null && data2[i, j] != "")
                                                {
                                                    sb.Append(GetLong(data2[i, j]));
                                                }
                                                else
                                                {
                                                    sb.Append("0");
                                                }
                                                break;
                                            case "Date":
                                                if (data2[i, j] != null && data2[i, j] != "")
                                                {
                                                    sb.Append(strSingleQuote);
                                                    sb.Append(data2[i, j]);
                                                    sb.Append(strSingleQuote);
                                                }
                                                else
                                                {
                                                    sb.Append("NULL");
                                                }
                                                break;
                                            default:
                                                break;
                                        }
                                        break;
                                }

                            }
                            sb.Append(")");
                            textInsertSQL += sb.ToString();
                        }
                    }

                    if (textInsertSQL != "")
                    {
                        cmd.CommandText = textInsertSQL;
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();
                    }

                    Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "単品入力項目が更新されました。 窓口ID = " + MadoguchiID, "MadoguchiUpdate_SQL", MadoguchiID);

                    // データセット
                    string TanpinJutakuDate = "NULL";
                    if (data[0, 1] != null && data[0, 1] != "")
                    {
                        TanpinJutakuDate = "'" + data[0, 1].ToString() + "'";
                    }

                    cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                                        " MadoguchiJutakubi         =  " + TanpinJutakuDate + " " +
                                        ",MadoguchiHachuubusho      = N'" + data[0, 4] + "' " +
                                        ",MadoguchiHachuuTantousha  = N'" + data[0, 6] + "' " +
                                        ",MadoguchiHachuuTEL        = N'" + data[0, 7] + "' " +
                                        ",MadoguchiHachuuFAX        = N'" + data[0, 8] + "' " +
                                        ",MadoguchiHachuuMail       = N'" + data[0, 9] + "' " +
                                        ",MadoguchiUpdateDate       = SYSDATETIME()" +
                                        ",MadoguchiUpdateUser       = N'" + UserInfos[0] + "' " +
                                        ",MadoguchiUpdateProgram    = '" + pgmName + methodName + "' " +
                                        " WHERE MadoguchiID = " + MadoguchiID;

                    Console.WriteLine(cmd.CommandText);
                    result = cmd.ExecuteNonQuery();

                    if (result <= 0)
                    {
                        mes += Environment.NewLine + "窓口情報への書き込みに失敗しました。";
                    }

                    mes += Environment.NewLine + GetMessage("I20603", "");

                }
                // 施工更新
                else if (tab == 7)
                {
                    //施工の画面モード取得
                    string sekouMode = data[0, 0];

                    //施工条件新規モードの場合
                    if ("0".Equals(sekouMode))
                    {
                        //採番No（SaibanNo）を取得
                        var dt = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                            "SaibanNo+SaibanCountupNo AS SaibanNo " +
                            "FROM " + "M_Saiban " +
                            "WHERE SaibanMei = 'SekouJoukenID' ";

                        //データ取得
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(dt);

                        string saibanSekouID = dt.Rows[0][0].ToString();

                        //採番No（SaibanNo）を更新
                        cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                            saibanSekouID + " WHERE SaibanMei = 'SekouJoukenID' ";

                        cmd.ExecuteNonQuery();

                        //施工条件（SekouJoukenu）テーブル登録
                        cmd.CommandText = "INSERT INTO SekouJouken( " +
                            "SekouJoukenID " +
                            ",SekouJoukenMeijishoID " +
                            ",SekouKoushuMei " +
                            ",SekouTenpuUmu " +
                            ",SekouGenbaHeimenzu " +
                            ",SekouDoshituKankeizu " +
                            ",SekouSuuryouKeisanzu " +
                            ",SekouHiruma " +
                            ",SekouYakan " +
                            ",SekouKiseiAri " +
                            ",SekouSagyouKouritsu " +
                            ",SekouKikai " +
                            ",SekouKasetu " +
                            ",SekouShizai " +
                            ",SekouKensetsu " +
                            ",SekouSuichuu " +
                            ",SekouSonota " +
                            ",SekouMemo1 " +
                            ",SekouMemo2 " +
                            ",MadoguchiID " +
                            ",SekouCreateDate " +
                            ",SekouCreateUser" +
                            ",SekouCreateProgram" +
                            ",SekouUpdateDate " +
                            ",SekouUpdateUser" +
                            ",SekouUpdateProgram" +
                            ",SekouDeleteFlag" +
                            ",SekouTenpuUmup1Ichizu01" +
                            ",SekouTenpuUmup1Sekou02" +
                            ",SekouTenpuUmup1Sankou03" +
                            ",SekouTenpuUmup1Ippan04" +
                            ",SekouTenpuUmup1Genba05" +
                            ",SekouTenpuUmup1Kako06" +
                            ",SekouTenpuUmup1Shousai07" +
                            ",SekouTenpuUmup1Doshitu08" +
                            ",SekouTenpuUmup1Sonota09" +
                            ",SekouTenpuUmup1Suuryou10" +
                            ",SekouTenpuUmup1Unpan11" +
                            ",SekouSekou2Rikujou01" +
                            ",SekouSekou2Suijou02" +
                            ",SekouSekou2Suichuu03" +
                            ",SekouSekou2Sonota04" +
                            ",SekouSekou3Tsuujou01" +
                            ",SekouSekou3Tsuujou02" +
                            ",SekouSekou3Sekou03" +
                            ",SekouSekou3Nihou04" +
                            ",SekouSekou3Sanpou05" +
                            ",SekouSagyou4Kankyou01" +
                            ",SekouSagyou4Sekou02" +
                            ",SekouSagyou4Joukuu03" +
                            ",SekouSagyou4Sonota04" +
                            ",SekouSagyou4Jinka05" +
                            ",SekouSagyou4Tokki06" +
                            ",SekouSagyou4Kankyou07" +
                            ",SekouSagyou5Koutusu01" +
                            ",SekouSagyou5Hannyuu02" +
                            ",SekouSagyou5Sonota03" +
                            ",SekouSagyou5Tokki04" +
                            ",SekouKasetsu6Shitei01" +
                            ",SekouKasetsu6Shitei02" +
                            ",SekouSekou7Shitei01" +
                            ",SekouSekou7Shitei02" +
                            ",SekouSonota8Shitei01" +
                            ",SekouSonota8Shitei02" +
                            ",SekouSonotaMemo03" +
                            " ) VALUES ( " +
                            saibanSekouID + " " +//採番テーブルから取得した値
                            ",N'" + ChangeSqlText(data[0, 1],0,0) + "' " + //施工条件明示書ID
                            ",N'" + ChangeSqlText(data[0, 2],0,0) + "' " + //工種名
                            "," + data[0, 3] + " " +//①施工計画書添付の有無
                            "," + data[0, 4] + " " +//②その他添付資料の現場平面図
                            "," + data[0, 5] + " " +//②その他添付資料の土質関係図
                            "," + data[0, 6] + " " +//②その他添付資料の数量計算書
                            "," + data[0, 7] + " " +//③施工時間帯指定の昼間
                            "," + data[0, 8] + " " + // ③施工時間帯指定の夜間
                            "," + data[0, 9] + " " + // ③施工時間帯指定の規制有り
                            "," + data[0, 10] + " " +//④施工条件他の作業効率
                            "," + data[0, 11] + " " + // ④施工条件他の施工機械の搬入経路
                            "," + data[0, 12] + " " +//④施工条件他の仮設条件
                            "," + data[0, 13] + " " +//④施工条件他の資材搬入
                            "," + data[0, 14] + " " +//⑤建設機械スペック指定
                            "," + data[0, 15] + " " +//⑥水中施行条件
                            "," + data[0, 16] + " " +//⑦その他
                            ",N'" + ChangeSqlText(data[0, 17],0,0) + "' " +//メモ1
                            ",N'" + ChangeSqlText(data[0, 18],0,0) + "' " +//メモ2
                            "," + data[0, 19] + " " +//窓口ID
                            ",SYSDATETIME() " +                             // 登録日時
                            ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                            ",'" + pgmName + methodName + "' " +            // 登録プログラム
                            ",SYSDATETIME() " +                             // 更新日時
                            ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                            ",'" + pgmName + methodName + "' " +            // 更新プログラム
                            ",0 " +                                         // 削除フラグ
                            "," + data[0, 23] + " " + // 3.添付資料の位置図
                            "," + data[0, 24] + " " + // 3.添付資料の施工計画書
                            "," + data[0, 25] + " " + // 3.添付資料の参考カタログ
                            "," + data[0, 26] + " " + // 3.添付資料の一般図・平面図
                            "," + data[0, 27] + " " + // 3.添付資料の現場写真
                            "," + data[0, 28] + " " + // 3.添付資料の過去報告書
                            "," + data[0, 29] + " " + // 3.添付資料の詳細図
                            "," + data[0, 30] + " " + // 3.添付資料の土質関係図（柱状図等）
                            "," + data[0, 31] + " " + // 3.添付資料のその他
                            "," + data[0, 32] + " " + // 3.添付資料の数量計算書
                            "," + data[0, 33] + " " + // 3.添付資料の運搬ルート図
                            "," + data[0, 34] + " " + // 5.(1)施工場所の陸上
                            "," + data[0, 35] + " " + // 5.(1)施工場所の水上
                            "," + data[0, 36] + " " + // 5.(1)施工場所の水中
                            "," + data[0, 37] + " " + // 5.(1)施工場所のその他
                            "," + data[0, 38] + " " + // 5.(2)施工時間帯の通常昼間施工（8:00~17:00）
                            "," + data[0, 39] + " " + // 5.(2)施工時間帯の通常夜間施工（20:00~5:00）
                            "," + data[0, 40] + " " + // 5.(2)施工時間帯の施工時間規制あり
                            "," + data[0, 41] + " " + // 5.(2)施工時間帯の二方施工（2交代制昼夜連続施工）
                            "," + data[0, 42] + " " + // 5.(2)施工時間帯の三方施工（3交代制24時間施工）
                            "," + data[0, 43] + " " + // 5.(3)作業環境の現場が狭隘
                            "," + data[0, 44] + " " + // 5.(3)作業環境の施工箇所が点在
                            "," + data[0, 45] + " " + // 5.(3)作業環境の上空制限あり
                            "," + data[0, 46] + " " + // 5.(3)作業環境のその他
                            "," + data[0, 47] + " " + // 5.(3)作業環境の人家に近接（近接施工）
                            "," + data[0, 48] + " " + // 5.(3)作業環境の特記すべき条件なし
                            "," + data[0, 49] + " " + // 5.(3)作業環境の環境対策あり（騒音・振動）
                            "," + data[0, 50] + " " + // 5.(4)施工機械・資材搬入経路の交通規制あり
                            "," + data[0, 51] + " " + // 5.(4)施工機械・資材搬入経路の搬入経路の制限（道路幅・時間など）
                            "," + data[0, 52] + " " + // 5.(4)施工機械・資材搬入経路のその他
                            "," + data[0, 53] + " " + // 5.(4)施工機械・資材搬入経路の特記すべき条件なし
                            "," + data[0, 54] + " " + // 5.(5)仮設条件の指定あり
                            "," + data[0, 55] + " " + // 5.(5)仮設条件の特記すべき条件なし
                            "," + data[0, 56] + " " + // 5.(6)施工機械スペック指定の指定あり
                            "," + data[0, 57] + " " + // 5.(6)施工機械スペック指定の指定なし
                            "," + data[0, 58] + " " + // 5.(7)その他条件の指定あり
                            "," + data[0, 59] + " " + // 5.(7)その他条件の特記すべき条件なし
                            ",N'" + ChangeSqlText(data[0, 60],0,0) + "' " + // メモ
                            " ) ";

                        cmd.ExecuteNonQuery();

                        //メッセージ表示
                        // I20705:施工条件を新規登録しました。
                        mes += GetMessage("I20705","");

                        Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "施工条件が新規作成されました。 窓口ID = " + MadoguchiID + " 明示書ID = " + ChangeSqlText(data[0, 1], 0, 0), "MadoguchiUpdate_SQL", MadoguchiID);

                    }
                    else
                    {
                        // Update
                        cmd.CommandText = "UPDATE SekouJouken SET " +
                        " SekouJoukenMeijishoID = " + "N'" + ChangeSqlText(data[0, 1], 0, 0) + "' " + //施工条件明示書ID
                        ",SekouKoushuMei  = " + "N'" + ChangeSqlText(data[0, 2], 0, 0) + "' " + //工種名
                        ",SekouTenpuUmu  = " + data[0, 3] + " " +//①施工計画書添付の有無
                        ",SekouGenbaHeimenzu  = " + data[0, 4] + " " +//②その他添付資料の現場平面図
                        ",SekouDoshituKankeizu  = " + data[0, 5] + " " +//②その他添付資料の土質関係図
                        ",SekouSuuryouKeisanzu  = " + data[0, 6] + " " +//②その他添付資料の数量計算書
                        ",SekouHiruma  = " + data[0, 7] + " " +//③施工時間帯指定の昼間
                        ",SekouYakan  = " + data[0, 8] + " " + // ③施工時間帯指定の夜間
                        ",SekouKiseiAri  = " + data[0, 9] + " " + // ③施工時間帯指定の規制有り
                        ",SekouSagyouKouritsu  = " + data[0, 10] + " " +//④施工条件他の作業効率
                        ",SekouKikai  = " + data[0, 11] + " " + // ④施工条件他の施工機械の搬入経路
                        ",SekouKasetu  = " + data[0, 12] + " " +//④施工条件他の仮設条件
                        ",SekouShizai  = " + data[0, 13] + " " +//④施工条件他の資材搬入
                        ",SekouKensetsu  = " + data[0, 14] + " " +//⑤建設機械スペック指定
                        ",SekouSuichuu  = " + data[0, 15] + " " +//⑥水中施行条件
                        ",SekouSonota  = " + data[0, 16] + " " +//⑦その他
                        ",SekouMemo1  = " + "N'" + ChangeSqlText(data[0, 17], 0, 0) + "' " +//メモ1
                        ",SekouMemo2  = " + "N'" + ChangeSqlText(data[0, 18], 0, 0) + "' " +//メモ2
                        ",SekouUpdateDate = SYSDATETIME() " +                     // 更新日時
                        ",SekouUpdateUser = N'" + UserInfos[0] + "' " +            // 更新ユーザ
                        ",SekouUpdateProgram = '" + pgmName + methodName + "' " + // 更新プログラム
                        ",SekouDeleteFlag = 0 " +                                 // 削除フラグ
                        ",SekouTenpuUmup1Ichizu01 = " + data[0, 23] + " " + // 3.添付資料の位置図
                        ",SekouTenpuUmup1Sekou02 = " + data[0, 24] + " " + // 3.添付資料の施工計画書
                        ",SekouTenpuUmup1Sankou03 = " + data[0, 25] + " " + // 3.添付資料の参考カタログ
                        ",SekouTenpuUmup1Ippan04 = " + data[0, 26] + " " + // 3.添付資料の一般図・平面図
                        ",SekouTenpuUmup1Genba05 = " + data[0, 27] + " " + // 3.添付資料の現場写真
                        ",SekouTenpuUmup1Kako06 = " + data[0, 28] + " " + // 3.添付資料の過去報告書
                        ",SekouTenpuUmup1Shousai07 = " + data[0, 29] + " " + // 3.添付資料の詳細図
                        ",SekouTenpuUmup1Doshitu08 = " + data[0, 30] + " " + // 3.添付資料の土質関係図（柱状図等）
                        ",SekouTenpuUmup1Sonota09 = " + data[0, 31] + " " + // 3.添付資料のその他
                        ",SekouTenpuUmup1Suuryou10 = " + data[0, 32] + " " + // 3.添付資料の数量計算書
                        ",SekouTenpuUmup1Unpan11 = " + data[0, 33] + " " + // 3.添付資料の運搬ルート図
                        ",SekouSekou2Rikujou01 = " + data[0, 34] + " " + // 5.(1)施工場所の陸上
                        ",SekouSekou2Suijou02 = " + data[0, 35] + " " + // 5.(1)施工場所の水上
                        ",SekouSekou2Suichuu03 = " + data[0, 36] + " " + // 5.(1)施工場所の水中
                        ",SekouSekou2Sonota04 = " + data[0, 37] + " " + // 5.(1)施工場所のその他
                        ",SekouSekou3Tsuujou01 = " + data[0, 38] + " " + // 5.(2)施工時間帯の通常昼間施工（8:00~17:00）
                        ",SekouSekou3Tsuujou02 = " + data[0, 39] + " " + // 5.(2)施工時間帯の通常夜間施工（20:00~5:00）
                        ",SekouSekou3Sekou03 = " + data[0, 40] + " " + // 5.(2)施工時間帯の施工時間規制あり
                        ",SekouSekou3Nihou04 = " + data[0, 41] + " " + // 5.(2)施工時間帯の二方施工（2交代制昼夜連続施工）
                        ",SekouSekou3Sanpou05 = " + data[0, 42] + " " + // 5.(2)施工時間帯の三方施工（3交代制24時間施工）
                        ",SekouSagyou4Kankyou01 = " + data[0, 43] + " " + // 5.(3)作業環境の現場が狭隘
                        ",SekouSagyou4Sekou02 = " + data[0, 44] + " " + // 5.(3)作業環境の施工箇所が点在
                        ",SekouSagyou4Joukuu03 = " + data[0, 45] + " " + // 5.(3)作業環境の上空制限あり
                        ",SekouSagyou4Sonota04 = " + data[0, 46] + " " + // 5.(3)作業環境のその他
                        ",SekouSagyou4Jinka05 = " + data[0, 47] + " " + // 5.(3)作業環境の人家に近接（近接施工）
                        ",SekouSagyou4Tokki06 = " + data[0, 48] + " " + // 5.(3)作業環境の特記すべき条件なし
                        ",SekouSagyou4Kankyou07 = " + data[0, 49] + " " + // 5.(3)作業環境の環境対策あり（騒音・振動）
                        ",SekouSagyou5Koutusu01 = " + data[0, 50] + " " + // 5.(4)施工機械・資材搬入経路の交通規制あり
                        ",SekouSagyou5Hannyuu02 = " + data[0, 51] + " " + // 5.(4)施工機械・資材搬入経路の搬入経路の制限（道路幅・時間など）
                        ",SekouSagyou5Sonota03 = " + data[0, 52] + " " + // 5.(4)施工機械・資材搬入経路のその他
                        ",SekouSagyou5Tokki04 = " + data[0, 53] + " " + // 5.(4)施工機械・資材搬入経路の特記すべき条件なし
                        ",SekouKasetsu6Shitei01 = " + data[0, 54] + " " + // 5.(5)仮設条件の指定あり
                        ",SekouKasetsu6Shitei02 = " + data[0, 55] + " " + // 5.(5)仮設条件の特記すべき条件なし
                        ",SekouSekou7Shitei01 = " + data[0, 56] + " " + // 5.(6)施工機械スペック指定の指定あり
                        ",SekouSekou7Shitei02 = " + data[0, 57] + " " + // 5.(6)施工機械スペック指定の指定なし
                        ",SekouSonota8Shitei01 = " + data[0, 58] + " " + // 5.(7)その他条件の指定あり
                        ",SekouSonota8Shitei02 = " + data[0, 59] + " " + // 5.(7)その他条件の特記すべき条件なし
                        ",SekouSonotaMemo03 = " + "N'" + ChangeSqlText(data[0, 60], 0, 0) + "' " + // メモ
                        "WHERE MadoguchiID = '" + data[0, 19] + "' AND SekouJoukenMeijishoID COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + ChangeSqlText(data[0, 1], 0, 0) + "' ";

                        cmd.ExecuteNonQuery();

                        //メッセージ表示
                        // I20706:施工条件を更新しました。
                        mes += GetMessage("I20706", "");

                        Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "施工条件が更新されました。 窓口ID = " + MadoguchiID + " 明示書ID = " + ChangeSqlText(data[0, 1], 0, 0), "MadoguchiUpdate_SQL", MadoguchiID);

                    }
                }

                // 応援受付状況タブ以外の場合
                if(tab != 5) 
                { 
                    transaction.Commit();

                    // 調査概要タブの更新 or 担当部所タブ
                    if((tab == 1 && !"insert".Equals(data[0, 0])) || tab == 2)
                    {
                        // 皇帝まもる連携
                        KouteiTantouBushoRenkei(MadoguchiID, UserInfos[0], UserInfos[2]);
                    }

                }
            }
            catch (ArithmeticException e)
            {
                transaction.Rollback();
                Console.WriteLine(e);
                return false;
            }
            finally
            {
                sqlconn.Close();
            }



            return true;
        }


        public Boolean MadoguchiUpdate_ErrorCheck(int tab, string[,] data, out string[] Errormes)
        {
            Boolean ErrorFlag = true;
            Errormes = null;

            if (tab == 6)
            {
                Errormes = new string[3];
                //if (data[0, 7] != null && data[0, 7] != "" && !System.Text.RegularExpressions.Regex.IsMatch(data[0, 7], @"^((0\d{1,4}-\d{1,4}-\d{4})|\+?(\d{10,12})|(\+\d{1,3}-\d{1,2}-\d{1,4}-\d{4})|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                if (data[0, 7] != null && data[0, 7] != "" && !System.Text.RegularExpressions.Regex.IsMatch(data[0, 7], @"^((0[0-9]{1,4}-[0-9]{1,4}-[0-9]{4})|\+?([0-9]{10,12})|(\+[0-9]{1,3}-[0-9]{1,2}-[0-9]{1,4}-[0-9]{4})|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    // E20603:電話番号を正しく入力してください。
                    Errormes[0] = GetMessage("E20603", "");
                    ErrorFlag = false;
                }
                //if (data[0, 8] != null && data[0, 8] != "" && !System.Text.RegularExpressions.Regex.IsMatch(data[0, 8], @"^((0\d{1,4}-\d{1,4}-\d{4})|\+?(\d{11,12})|(\+\d{1,3}-\d{1,2}-\d{1,4}-\d{4})|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                if (data[0, 8] != null && data[0, 8] != "" && !System.Text.RegularExpressions.Regex.IsMatch(data[0, 8], @"^((0[0-9]{1,4}-[0-9]{1,4}-[0-9]{4})|\+?([0-9]{10,12})|(\+[0-9]{1,3}-[0-9]{1,2}-[0-9]{1,4}-[0-9]{4})|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    // E20604:FAX番号を正しく入力してください。
                    Errormes[1] = GetMessage("E20604", "");
                    ErrorFlag = false;
                }
                //if (data[0, 9] != null && data[0, 9] != "" && !System.Text.RegularExpressions.Regex.IsMatch(data[0, 9], @"^((\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                if (data[0, 9] != null && data[0, 9] != "" && !System.Text.RegularExpressions.Regex.IsMatch(data[0, 9], @"^(([a-zA-Z_0-9]+([-+.'][a-zA-Z_0-9]+)*@[a-zA-Z_0-9]+([-.][a-zA-Z_0-9]+)*\.[a-zA-Z_0-9]+([-.][a-zA-Z_0-9]+)*)|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    // E20605:メールアドレスを正しく入力してください。
                    Errormes[2] = GetMessage("E20605", "");
                    ErrorFlag = false;
                }
            }

            return ErrorFlag;
        }



        // 窓口ミハル調査概要タブの調査品目取込後の担当者更新処理
        public Boolean MadoguchiHinmokuRenkeiUpdate_SQL(string MadoguchiID, string gamenMode, string UpdateKojinCD, out string mes)
        {
            // gamenMode
            // Madoguchi：窓口ミハル Tokumei：特命課長 Jibun：自分大臣

            string methodName = ".MadoguchiHinmokuRenkeiUpdate_SQL";

            using (var conn = new SqlConnection(connStr))
            {
                mes = "";
                conn.Open();

                var cmd = conn.CreateCommand();
                SqlTransaction transaction = conn.BeginTransaction();
                cmd.Transaction = transaction;

                var MagoguchiL1ChousaCDList = new List<string>();
                var MadoguchiL1ChousaBushoCDList = new List<string>();
                var MadoguchiL1ChousaBushoList = new List<string>();
                var MadoguchiL1ChousaRyakumeiList = new List<string>();
                var MadoguchiL1ChousaTantoushaCDList = new List<string>();
                var MadoguchiL1ChousaTantoushaList = new List<string>();
                var MadoguchiL1ChousaShimekiribiList = new List<string>();
                var MadoguchiL1ChousaShinchokuList = new List<string>();
                // 新特調での追加項目
                var MadoguchiL1ShuTantouFlagList = new List<string>();
                var MadoguchiL1ShuTantouHinmokusuList = new List<string>();
                var MadoguchiL1FukuTantouFlagList = new List<string>();
                var MadoguchiL1FukuTantouHinmokusuList = new List<string>();

                // 調査担当者判別用 0:調査担当者 1:副調査担当者1 2:副調査担当者2
                var MadoguchiL1ChousaKindList = new List<string>();

                // 親データ（特調番号、枝番、集計表、登録年度、発注者名・課名、調査担当者への締切日、登録日）
                string TokuchoBangou = "";
                string UketsukeBangou = "";
                string UketsukeBangouEdaban = "";
                string ShukeihyoLink = "";
                string TourokuNendo = "";
                string HachuuKikanmei = "";
                string Shimekiribi = "";
                string Tourokubi = "";
                string ShiryouHolder = "";
                
                //// 部所マスタ検索用
                //string mb_discript_Kamei = "Mst_Busho.Kamei ";
                //string mb_discript_kanriboKamei = "Mst_Busho.BushokanriboKamei ";
                //string mb_value = "Mst_Busho.GyoumuBushoCD ";
                //string mb_table = "Mst_Busho";
                //string mb_where = "GyoumuBushoCD != '999990' AND GyoumuBushoCD != '127900' AND BushoNewOld <= 1 AND BushoMadoguchiHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 ";
                //DataTable mb_DataTable = new DataTable();

                //// 調査員マスタ検索用
                //string mc_discript = "Mst_Chousain.ChousainMei ";
                //string mc_value = "Mst_Chousain.KojinCD ";
                //string mc_table = "Mst_Chousain";
                //string mc_where = "RetireFLG = 1 AND ISNULL(ChousainDeleteFlag,0) = 0 ";
                //DataTable mc_DataTable = new DataTable();

                // 品目数カウント検索用
                string cnt_discript = "''";
                string cnt_value = "count(*) ";
                string cnt_table = "ChousaHinmoku with(nolock) "; // Lockせずに取得

                // 検索時に条件に付ける
                // ▼調査担当者の品目数
                // HinmokuChousainCD = ''
                // ▼副調査担当者1または副調査担当者2
                // (HinmokuFukuChousainCD1 = '' or HinmokuFukuChousainCD2 = '')
                string cnt_where = "ISNULL(ChousaDeleteFlag,0) = 0 AND MadoguchiID = '" + MadoguchiID + "' AND ";
                DataTable cnt_DataTable = new DataTable();
                // カウントを入れる用
                string cnt_str = "";

                // HinmokuRyakuBushoCD + ',' +HinmokuChousainCD を入れる
                var i_SDTMadoguchiL1Chousa_p = new List<string>();

                int num = 0;
                DateTime dateTime1 = DateTime.Now;
                DateTime dateTime2 = DateTime.Now;
                String dateFormat = "yyyy/MM/dd";

                int shinchokuJoukyou1 = 0;
                int shinchokuJoukyou2 = 0;

                // 画面表示メッセージ判断のフラグ
                int updmessage1 = 0;
                int updmessage2 = 0;
                int updmessage3 = 0;

                try
                {
                    // 親の情報を取得
                    cmd.CommandText = "SELECT  " +
                        "MadoguchiUketsukeBangou " +          // 0:特調番号
                        ",MadoguchiUketsukeBangouEdaban " +   // 1:特調番号枝番
                        ",MadoguchiShukeiHyoFolder " +        // 2:集計表
                        ",MadoguchiTourokuNendo " +           // 3:登録年度
                        ",MadoguchiHachuuKikanmei " +         // 4:発注者詳細名
                        ",MadoguchiShimekiribi " +            // 5:調査担当者への締切日
                        ",MadoguchiTourokubi " +              // 6:登録日
                                                              // No1719 MadoguchiJouhouMadoguchiL1Chouの[MadoguchiL1SiryouHolder]が特調奉行、工程まもる側に登録されない。
                        ",ISNULL(MadoguchiShiryouHolder,'') MadoguchiShiryouHolder " +          // 7:資料フォルダ
                        "FROM MadoguchiJouhou  mj " +
                        "WHERE MadoguchiID = '" + MadoguchiID + "' ";

                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    var dtMadoguchiL1 = new DataTable();
                    sda.Fill(dtMadoguchiL1);

                    // 取得データがある場合
                    if (dtMadoguchiL1.Rows.Count > 0)
                    {
                        UketsukeBangou = dtMadoguchiL1.Rows[0][0].ToString();
                        UketsukeBangouEdaban = dtMadoguchiL1.Rows[0][1].ToString();
                        ShukeihyoLink = dtMadoguchiL1.Rows[0][2].ToString();
                        TourokuNendo = dtMadoguchiL1.Rows[0][3].ToString();
                        HachuuKikanmei = dtMadoguchiL1.Rows[0][4].ToString();
                        TokuchoBangou = UketsukeBangou + "-" + UketsukeBangouEdaban;
                        Shimekiribi = dtMadoguchiL1.Rows[0][5].ToString();
                        Tourokubi = dtMadoguchiL1.Rows[0][6].ToString();
                        ShiryouHolder = dtMadoguchiL1.Rows[0][7].ToString();
                    }

                    // 調査担当者
                    cmd.CommandText = "SELECT distinct " +
                    "0 as  ChousaCD " +                 // 0:
                    ",HinmokuRyakuBushoCD " +           // 1:部所CD   HinmokuRyakuBushoCD
                    ",mb.BushoKanriboKamei " +          // 2:課名     BushoKanriboKamei
                    ",mb.BushokanriboKameiRaku " +      // 3:部所略名 BushokanriboKameiRaku
                    ",ch.HinmokuChousainCD  " +         // 4:調査員CD HinmokuChousainCD
                    ",mc.ChousainMei  " +               // 5:調査員名 ChousainMei
                    ",FORMAT(ChousaHinmokuShimekiribi,'yyyy/MM/dd') AS  Shimekiribi" +     // 6:締切日   ChousaHinmokuShimekiribi
                    ",ChousaShinchokuJoukyou  " +       // 7:進捗状況 ChousaShinchokuJoukyou
                    "FROM ChousaHinmoku  ch " +
                    "LEFT JOIN Mst_Busho mb on mb.GyoumuBushoCD = ch.HinmokuRyakuBushoCD " +
                    "LEFT JOIN Mst_Chousain mc on mc.KojinCD = ch.HinmokuChousainCD " +
                    "WHERE MadoguchiID = '" + MadoguchiID + "' " +
                    "AND ch.HinmokuRyakuBushoCD is not null  " +
                    "AND ch.HinmokuRyakuBushoCD != '' " +
                    "AND HinmokuRyakuBushoCD != '0' " +
                    "order by ch.HinmokuRyakuBushoCD,HinmokuChousainCD ";

                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    dtMadoguchiL1 = new DataTable();
                    sda.Fill(dtMadoguchiL1);

                    // 取得データがある場合
                    if (dtMadoguchiL1.Rows.Count > 0)
                    {
                        for (int i = 0; dtMadoguchiL1.Rows.Count > i; i++)
                        {
                            // 
                            if (i_SDTMadoguchiL1Chousa_p.IndexOf(dtMadoguchiL1.Rows[i][1].ToString() + "," + dtMadoguchiL1.Rows[i][4].ToString()) > -1)
                            {
                                // 見つかった場合
                                num = i_SDTMadoguchiL1Chousa_p.IndexOf(dtMadoguchiL1.Rows[i][1].ToString() + "," + dtMadoguchiL1.Rows[i][4].ToString());

                                dateTime1 = DateTime.ParseExact(MadoguchiL1ChousaShimekiribiList[num], dateFormat, null);
                                dateTime2 = DateTime.ParseExact(dtMadoguchiL1.Rows[i][6].ToString(), dateFormat, null);

                                // 締切日を比較し、締切日が小さい方を詰めなおす
                                if (dateTime1 > dateTime2)
                                {
                                    MadoguchiL1ChousaShimekiribiList[num] = dtMadoguchiL1.Rows[i][6].ToString();
                                }
                                shinchokuJoukyou1 = 0;
                                shinchokuJoukyou2 = 0;

                                int.TryParse(MadoguchiL1ChousaShinchokuList[num], out shinchokuJoukyou1);
                                int.TryParse(dtMadoguchiL1.Rows[i][7].ToString(), out shinchokuJoukyou2);
                                // 進捗状況を比較し、進捗状況が小さい方を詰めなおす
                                if (shinchokuJoukyou1 > shinchokuJoukyou2)
                                {
                                    MadoguchiL1ChousaShinchokuList[num] = dtMadoguchiL1.Rows[i][7].ToString();
                                }

                            }
                            else
                            {
                                // 見つからなかった場合
                                // HinmokuRyakuBushoCD + ',' +HinmokuChousainCD を入れる
                                i_SDTMadoguchiL1Chousa_p.Add(dtMadoguchiL1.Rows[i][1].ToString() + "," + dtMadoguchiL1.Rows[i][4].ToString());

                                // Listに詰める
                                MagoguchiL1ChousaCDList.Add(dtMadoguchiL1.Rows[i][0].ToString());
                                MadoguchiL1ChousaBushoCDList.Add(dtMadoguchiL1.Rows[i][1].ToString());
                                MadoguchiL1ChousaBushoList.Add(dtMadoguchiL1.Rows[i][2].ToString());
                                MadoguchiL1ChousaRyakumeiList.Add(dtMadoguchiL1.Rows[i][3].ToString());
                                MadoguchiL1ChousaTantoushaCDList.Add(dtMadoguchiL1.Rows[i][4].ToString());
                                MadoguchiL1ChousaTantoushaList.Add(dtMadoguchiL1.Rows[i][5].ToString());
                                MadoguchiL1ChousaShimekiribiList.Add(dtMadoguchiL1.Rows[i][6].ToString());
                                MadoguchiL1ChousaShinchokuList.Add(dtMadoguchiL1.Rows[i][7].ToString());

                                cnt_str = "0";
                                // 調査員CDが空ではない場合、品目数を取得する
                                if (dtMadoguchiL1.Rows[i][4].ToString() != "")
                                {
                                    cnt_DataTable = new DataTable();
                                    //cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " HinmokuChousainCD = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' ");
                                    cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " HinmokuChousainCD = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' AND HinmokuRyakuBushoCD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "' ");
                                    if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                    {
                                        cnt_str = cnt_DataTable.Rows[0][0].ToString();
                                    }
                                }
                                else
                                {
                                    cnt_DataTable = new DataTable();
                                    //cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " HinmokuChousainCD = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' ");
                                    cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " HinmokuChousainCD is null AND HinmokuRyakuBushoCD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "' ");
                                    if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                    {
                                        cnt_str = cnt_DataTable.Rows[0][0].ToString();
                                    }
                                }

                                MadoguchiL1ShuTantouFlagList.Add("1");
                                MadoguchiL1ShuTantouHinmokusuList.Add(cnt_str);
                                MadoguchiL1FukuTantouFlagList.Add("0");
                                MadoguchiL1FukuTantouHinmokusuList.Add("0");

                                // 調査担当者判別用 0:調査担当者 1:副調査担当者1 2:副調査担当者2
                                MadoguchiL1ChousaKindList.Add("0");

                                // debugログに部所略名 調査員名を書き込み
                                outputLogger("MadoguchiHinmokuRenkeiUpdate_SQL", "担当者追加 " + dtMadoguchiL1.Rows[i][3].ToString() + " " + dtMadoguchiL1.Rows[i][5].ToString(), "insert", "DEBUG");
                            }
                        }
                    }

                    // 副調査担当者1
                    cmd.CommandText = "SELECT distinct " +
                        "0 as  ChousaCD " +                 // 0:
                        ",HinmokuRyakuBushoFuku1CD " +      // 1:部所CD   HinmokuRyakuBushoFuku1CD
                        ",mb.BushoKanriboKamei " +          // 2:課名     BushoKanriboKamei
                        ",mb.BushokanriboKameiRaku " +      // 3:部所略名 BushokanriboKameiRaku
                        ",ch.HinmokuFukuChousainCD1  " +    // 4:調査員CD HinmokuFukuChousainCD1
                        ",mc.ChousainMei  " +               // 5:調査員名 ChousainMei
                        ",FORMAT(ChousaHinmokuShimekiribi,'yyyy/MM/dd') AS Shimekiribi " +     // 6:締切日   ChousaHinmokuShimekiribi
                        ",ChousaShinchokuJoukyou  " +       // 7:進捗状況 ChousaShinchokuJoukyou
                        "FROM ChousaHinmoku ch " +
                        "LEFT JOIN Mst_Busho mb on mb.GyoumuBushoCD = ch.HinmokuRyakuBushoFuku1CD " +
                        "LEFT JOIN Mst_Chousain mc on mc.KojinCD = ch.HinmokuFukuChousainCD1 " +
                        "WHERE MadoguchiID = '" + MadoguchiID + "' " +
                        "AND ch.HinmokuRyakuBushoFuku1CD is not null  " +
                        "AND ch.HinmokuRyakuBushoFuku1CD != '' " +
                        "AND HinmokuRyakuBushoFuku1CD != '0' " +
                        //不具合No1356(1124)　下記の条件があると担当部署に登録されないためコメントアウト
                        //"AND ch.HinmokuFukuChousainCD1 is not null " + // 998 副担当者1,2未割当は除外
                        "order by ch.HinmokuRyakuBushoFuku1CD,HinmokuFukuChousainCD1 ";

                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    dtMadoguchiL1 = new DataTable();
                    sda.Fill(dtMadoguchiL1);

                    // 取得データがある場合
                    if (dtMadoguchiL1.Rows.Count > 0)
                    {
                        for (int i = 0; dtMadoguchiL1.Rows.Count > i; i++)
                        {
                            // 
                            if (i_SDTMadoguchiL1Chousa_p.IndexOf(dtMadoguchiL1.Rows[i][1].ToString() + "," + dtMadoguchiL1.Rows[i][4].ToString()) > -1)
                            {
                                // 見つかった場合
                                num = i_SDTMadoguchiL1Chousa_p.IndexOf(dtMadoguchiL1.Rows[i][1].ToString() + "," + dtMadoguchiL1.Rows[i][4].ToString());

                                dateTime1 = DateTime.ParseExact(MadoguchiL1ChousaShimekiribiList[num], dateFormat, null);
                                dateTime2 = DateTime.ParseExact(dtMadoguchiL1.Rows[i][6].ToString(), dateFormat, null);

                                // 締切日を比較し、締切日が小さい方を詰めなおす
                                if (dateTime1 > dateTime2)
                                {
                                    MadoguchiL1ChousaShimekiribiList[num] = dtMadoguchiL1.Rows[i][6].ToString();
                                }
                                shinchokuJoukyou1 = 0;
                                shinchokuJoukyou2 = 0;

                                int.TryParse(MadoguchiL1ChousaShinchokuList[num], out shinchokuJoukyou1);
                                int.TryParse(dtMadoguchiL1.Rows[i][7].ToString(), out shinchokuJoukyou2);
                                // 進捗状況を比較し、進捗状況が小さい方を詰めなおす
                                if (shinchokuJoukyou1 > shinchokuJoukyou2)
                                {
                                    MadoguchiL1ChousaShinchokuList[num] = dtMadoguchiL1.Rows[i][7].ToString();
                                }

                                cnt_str = "0";
                                // 調査員CDが空ではない場合、品目数を取得する
                                if (dtMadoguchiL1.Rows[i][4].ToString() != "")
                                {
                                    cnt_DataTable = new DataTable();
                                    //cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " (HinmokuFukuChousainCD1 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' or HinmokuFukuChousainCD2 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "') ");
                                    cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " ((HinmokuFukuChousainCD1 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' AND HinmokuRyakuBushoFuku1CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "') or (HinmokuFukuChousainCD2 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' AND HinmokuRyakuBushoFuku2CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "')) ");
                                    if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                    {
                                        cnt_str = cnt_DataTable.Rows[0][0].ToString();
                                    }
                                }
                                else
                                {
                                    cnt_DataTable = new DataTable();
                                    //cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " (HinmokuFukuChousainCD1 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' or HinmokuFukuChousainCD2 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "') ");
                                    cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " ((HinmokuFukuChousainCD1 is null AND HinmokuRyakuBushoFuku1CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "') or (HinmokuFukuChousainCD2 is null AND HinmokuRyakuBushoFuku2CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "')) ");
                                    if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                    {
                                        cnt_str = cnt_DataTable.Rows[0][0].ToString();
                                    }
                                }

                                MadoguchiL1FukuTantouFlagList.Add("1");
                                MadoguchiL1FukuTantouHinmokusuList.Add(cnt_str);

                            }
                            else
                            {
                                // 見つからなかった場合
                                // HinmokuRyakuBushoCD + ',' +HinmokuChousainCD を入れる
                                i_SDTMadoguchiL1Chousa_p.Add(dtMadoguchiL1.Rows[i][1].ToString() + "," + dtMadoguchiL1.Rows[i][4].ToString());

                                // Listに詰める
                                MagoguchiL1ChousaCDList.Add(dtMadoguchiL1.Rows[i][0].ToString());
                                MadoguchiL1ChousaBushoCDList.Add(dtMadoguchiL1.Rows[i][1].ToString());
                                MadoguchiL1ChousaBushoList.Add(dtMadoguchiL1.Rows[i][2].ToString());
                                MadoguchiL1ChousaRyakumeiList.Add(dtMadoguchiL1.Rows[i][3].ToString());
                                MadoguchiL1ChousaTantoushaCDList.Add(dtMadoguchiL1.Rows[i][4].ToString());
                                MadoguchiL1ChousaTantoushaList.Add(dtMadoguchiL1.Rows[i][5].ToString());
                                MadoguchiL1ChousaShimekiribiList.Add(dtMadoguchiL1.Rows[i][6].ToString());
                                MadoguchiL1ChousaShinchokuList.Add(dtMadoguchiL1.Rows[i][7].ToString());

                                cnt_str = "0";
                                // 調査員CDが空ではない場合、品目数を取得する
                                if (dtMadoguchiL1.Rows[i][4].ToString() != "")
                                {
                                    cnt_DataTable = new DataTable();
                                    //cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " (HinmokuFukuChousainCD1 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' or HinmokuFukuChousainCD2 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "') ");
                                    cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " ((HinmokuFukuChousainCD1 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' AND HinmokuRyakuBushoFuku1CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "') or (HinmokuFukuChousainCD2 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' AND HinmokuRyakuBushoFuku2CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "')) ");
                                    if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                    {
                                        cnt_str = cnt_DataTable.Rows[0][0].ToString();
                                    }
                                }
                                else
                                {
                                    cnt_DataTable = new DataTable();
                                    cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " ((HinmokuFukuChousainCD1 is null AND HinmokuRyakuBushoFuku1CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "') or (HinmokuFukuChousainCD2 is null AND HinmokuRyakuBushoFuku2CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "')) ");
                                    if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                    {
                                        cnt_str = cnt_DataTable.Rows[0][0].ToString();
                                    }
                                }

                                MadoguchiL1ShuTantouFlagList.Add("0");
                                MadoguchiL1ShuTantouHinmokusuList.Add("0");
                                MadoguchiL1FukuTantouFlagList.Add("1");
                                MadoguchiL1FukuTantouHinmokusuList.Add(cnt_str);

                                // 調査担当者判別用 0:調査担当者 1:副調査担当者1 2:副調査担当者2
                                MadoguchiL1ChousaKindList.Add("1");

                                // debugログに部所略名 調査員名を書き込み
                                outputLogger("MadoguchiHinmokuRenkeiUpdate_SQL", "副担当者1追加 " + dtMadoguchiL1.Rows[i][3].ToString() + " " + dtMadoguchiL1.Rows[i][5].ToString(), "insert", "DEBUG");
                            }
                        }
                    }

                    // 副調査担当者2
                    cmd.CommandText = "SELECT distinct " +
                        "0 as  ChousaCD " +                 // 0:
                        ",HinmokuRyakuBushoFuku2CD " +      // 1:部所CD   HinmokuRyakuBushoFuku2CD
                        ",mb.BushoKanriboKamei " +          // 2:課名     BushoKanriboKamei
                        ",mb.BushokanriboKameiRaku " +      // 3:部所略名 BushokanriboKameiRaku
                        ",ch.HinmokuFukuChousainCD2  " +    // 4:調査員CD HinmokuFukuChousainCD2
                        ",mc.ChousainMei  " +               // 5:調査員名 ChousainMei
                        ",FORMAT(ChousaHinmokuShimekiribi,'yyyy/MM/dd') AS  Shimekiribi " +     // 6:締切日   ChousaHinmokuShimekiribi
                        ",ChousaShinchokuJoukyou  " +       // 7:進捗状況 ChousaShinchokuJoukyou
                        "FROM ChousaHinmoku  ch " +
                        "LEFT JOIN Mst_Busho mb on mb.GyoumuBushoCD = ch.HinmokuRyakuBushoFuku2CD " +
                        "LEFT JOIN Mst_Chousain mc on mc.KojinCD = ch.HinmokuFukuChousainCD2 " +
                        "WHERE MadoguchiID = '" + MadoguchiID + "' " +
                        "AND ch.HinmokuRyakuBushoFuku2CD is not null  " +
                        "AND ch.HinmokuRyakuBushoFuku2CD != '' " +
                        "AND HinmokuRyakuBushoFuku2CD != '0' " +
                        //不具合No1356(1124)　下記の条件があると担当部署に登録されないためコメントアウト
                        //"AND ch.HinmokuFukuChousainCD2 is not null " + // 998 副担当者1,2未割当は除外
                        "order by ch.HinmokuRyakuBushoFuku2CD,HinmokuFukuChousainCD2 ";

                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    dtMadoguchiL1 = new DataTable();
                    sda.Fill(dtMadoguchiL1);

                    // 取得データがある場合
                    if (dtMadoguchiL1.Rows.Count > 0)
                    {
                        for (int i = 0; dtMadoguchiL1.Rows.Count > i; i++)
                        {
                            // 
                            if (i_SDTMadoguchiL1Chousa_p.IndexOf(dtMadoguchiL1.Rows[i][1].ToString() + "," + dtMadoguchiL1.Rows[i][4].ToString()) > -1)
                            {
                                // 見つかった場合
                                num = i_SDTMadoguchiL1Chousa_p.IndexOf(dtMadoguchiL1.Rows[i][1].ToString() + "," + dtMadoguchiL1.Rows[i][4].ToString());

                                dateTime1 = DateTime.ParseExact(MadoguchiL1ChousaShimekiribiList[num], dateFormat, null);
                                dateTime2 = DateTime.ParseExact(dtMadoguchiL1.Rows[i][6].ToString(), dateFormat, null);

                                // 締切日を比較し、締切日が小さい方を詰めなおす
                                if (dateTime1 > dateTime2)
                                {
                                    MadoguchiL1ChousaShimekiribiList[num] = dtMadoguchiL1.Rows[i][6].ToString();
                                }
                                shinchokuJoukyou1 = 0;
                                shinchokuJoukyou2 = 0;

                                int.TryParse(MadoguchiL1ChousaShinchokuList[num], out shinchokuJoukyou1);
                                int.TryParse(dtMadoguchiL1.Rows[i][7].ToString(), out shinchokuJoukyou2);
                                // 進捗状況を比較し、進捗状況が小さい方を詰めなおす
                                if (shinchokuJoukyou1 > shinchokuJoukyou2)
                                {
                                    MadoguchiL1ChousaShinchokuList[num] = dtMadoguchiL1.Rows[i][7].ToString();
                                }
                                cnt_str = "0";
                                // 調査員CDが空ではない場合、品目数を取得する
                                if (dtMadoguchiL1.Rows[i][4].ToString() != "")
                                {
                                    cnt_DataTable = new DataTable();
                                    //cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " (HinmokuFukuChousainCD1 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' or HinmokuFukuChousainCD2 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "') ");
                                    cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " ((HinmokuFukuChousainCD1 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' AND HinmokuRyakuBushoFuku1CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "') or (HinmokuFukuChousainCD2 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' AND HinmokuRyakuBushoFuku2CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "')) ");
                                    if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                    {
                                        cnt_str = cnt_DataTable.Rows[0][0].ToString();
                                    }
                                }
                                else
                                {
                                    cnt_DataTable = new DataTable();
                                    cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " ((HinmokuFukuChousainCD1 is null AND HinmokuRyakuBushoFuku1CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "') or (HinmokuFukuChousainCD2 is null AND HinmokuRyakuBushoFuku2CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "')) ");
                                    if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                    {
                                        cnt_str = cnt_DataTable.Rows[0][0].ToString();
                                    }
                                }

                                MadoguchiL1FukuTantouFlagList.Add("1");
                                MadoguchiL1FukuTantouHinmokusuList.Add(cnt_str);
                            }
                            else
                            {
                                // 見つからなかった場合
                                // HinmokuRyakuBushoCD + ',' +HinmokuChousainCD を入れる
                                i_SDTMadoguchiL1Chousa_p.Add(dtMadoguchiL1.Rows[i][1].ToString() + "," + dtMadoguchiL1.Rows[i][4].ToString());

                                // Listに詰める
                                MagoguchiL1ChousaCDList.Add(dtMadoguchiL1.Rows[i][0].ToString());
                                MadoguchiL1ChousaBushoCDList.Add(dtMadoguchiL1.Rows[i][1].ToString());
                                MadoguchiL1ChousaBushoList.Add(dtMadoguchiL1.Rows[i][2].ToString());
                                MadoguchiL1ChousaRyakumeiList.Add(dtMadoguchiL1.Rows[i][3].ToString());
                                MadoguchiL1ChousaTantoushaCDList.Add(dtMadoguchiL1.Rows[i][4].ToString());
                                MadoguchiL1ChousaTantoushaList.Add(dtMadoguchiL1.Rows[i][5].ToString());
                                MadoguchiL1ChousaShimekiribiList.Add(dtMadoguchiL1.Rows[i][6].ToString());
                                MadoguchiL1ChousaShinchokuList.Add(dtMadoguchiL1.Rows[i][7].ToString());

                                cnt_str = "0";
                                // 調査員CDが空ではない場合、品目数を取得する
                                if (dtMadoguchiL1.Rows[i][4].ToString() != "")
                                {
                                    cnt_DataTable = new DataTable();
                                    //cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " (HinmokuFukuChousainCD1 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' or HinmokuFukuChousainCD2 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "') ");
                                    cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " ((HinmokuFukuChousainCD1 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' AND HinmokuRyakuBushoFuku1CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "') or (HinmokuFukuChousainCD2 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' AND HinmokuRyakuBushoFuku2CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "')) ");
                                    if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                    {
                                        cnt_str = cnt_DataTable.Rows[0][0].ToString();
                                    }
                                }
                                else
                                {
                                    cnt_DataTable = new DataTable();
                                    //cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " (HinmokuFukuChousainCD1 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "' or HinmokuFukuChousainCD2 = '" + dtMadoguchiL1.Rows[i][4].ToString() + "') ");
                                    cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " ((HinmokuFukuChousainCD1 is null AND HinmokuRyakuBushoFuku1CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "') or (HinmokuFukuChousainCD2 is null AND HinmokuRyakuBushoFuku2CD = '" + dtMadoguchiL1.Rows[i][1].ToString() + "')) ");
                                    if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                    {
                                        cnt_str = cnt_DataTable.Rows[0][0].ToString();
                                    }
                                }

                                MadoguchiL1ShuTantouFlagList.Add("0");
                                MadoguchiL1ShuTantouHinmokusuList.Add("0");
                                MadoguchiL1FukuTantouFlagList.Add("1");
                                MadoguchiL1FukuTantouHinmokusuList.Add(cnt_str);

                                // 調査担当者判別用 0:調査担当者 1:副調査担当者1 2:副調査担当者2
                                MadoguchiL1ChousaKindList.Add("2");

                                // debugログに部所略名 調査員名を書き込み
                                outputLogger("MadoguchiHinmokuRenkeiUpdate_SQL", "副担当者2追加 " + dtMadoguchiL1.Rows[i][3].ToString() + " " + dtMadoguchiL1.Rows[i][5].ToString(), "insert", "DEBUG");
                            }
                        }
                    }


                    var MagoguchiL1ChousaCDList_c = new List<string>();
                    var MadoguchiL1ChousaBushoCDList_c = new List<string>();
                    var MadoguchiL1ChousaBushoList_c = new List<string>();
                    var MadoguchiL1ChousaRyakumeiList_c = new List<string>();
                    var MadoguchiL1ChousaTantoushaCDList_c = new List<string>();
                    var MadoguchiL1ChousaTantoushaList_c = new List<string>();
                    var MadoguchiL1ChousaShimekiribiList_c = new List<string>();
                    var MadoguchiL1ChousaShinchokuList_c = new List<string>();
                    var MadoguchiL1ChousaKakuninList_c = new List<string>();
                    var MadoguchiL1MemoList_c = new List<string>();
                    var MadoguchiL1BunyaList_c = new List<string>();
                    var MadoguchiL1BunruiList_c = new List<string>();

                    // 新での追加項目
                    var MadoguchiL1TokuchoBangou_c = new List<string>();
                    var MadoguchiL1UketsukeBangou_c = new List<string>();
                    var MadoguchiL1UketsukeBangouEdaban_c = new List<string>();
                    var MadoguchiL1ShuTantouFlag_c = new List<string>();
                    var MadoguchiL1ShuTantouHinmokusu_c = new List<string>();
                    var MadoguchiL1FukuTantouFlag_c = new List<string>();
                    var MadoguchiL1FukuTantouHinmokusu_c = new List<string>();
                    var MadoguchiL1ShukeihyoLink_c = new List<string>();

                    var i_keyList_c = new List<string>();
                    var i_modeList_c = new List<string>();
                    var i_shufukuList_c = new List<string>();

                    // 今既に窓口子テーブルに存在しているデータを取得
                    //cmd.CommandText = "SELECT " +
                    //     "MadoguchiL1ChousaCD " +                   // 0:MagoguchiL1ChousaCD
                    //     ",MadoguchiL1ChousaBushoCD " +             // 1:部所CD   MadoguchiL1ChousaBushoCD
                    //     ",MadoguchiL1ChousaBusho " +               // 2:部所名   MadoguchiL1ChousaBusho
                    //     ",MadoguchiL1ChousaRyakumei " +            // 3:部所略名 MadoguchiL1ChousaRyakumei
                    //     ",MadoguchiL1ChousaTantoushaCD  " +        // 4:調査員CD MadoguchiL1ChousaTantoushaCD
                    //     ",MadoguchiL1ChousaTantousha  " +          // 5:調査員名 MadoguchiL1ChousaTantousha
                    //     ",FORMAT(MadoguchiL1ChousaShimekiribi,'yyyy/MM/dd') AS  Shimekiribi " +     // 6:締切日   MadoguchiL1ChousaShimekiribi
                    //     ",MadoguchiL1ChousaShinchoku  " +          // 7:進捗状況 MadoguchiL1ChousaShinchoku
                    //     ",MadoguchiL1ChousaKakunin  " +            // 8:確認     MadoguchiL1ChousaKakunin
                    //     ",MadoguchiL1Memo  " +                     // 9:メモ     MadoguchiL1Memo
                    //     ",MadoguchiL1Bunya  " +                    // 10:分野    MadoguchiL1Bunya
                    //     ",MadoguchiL1Bunrui  " +                   // 11:分類    MadoguchiL1Bunrui
                    //     ",'" + MadoguchiID + "' AS MadoguchiID " + // 12:窓口ID  MadoguchiID
                    //     "FROM MadoguchiJouhouMadoguchiL1Chou " +
                    //     "WHERE MadoguchiID = '" + MadoguchiID + "' " +
                    //     "order by MadoguchiID,MadoguchiL1ChousaCD ";

                    cmd.CommandText = "SELECT " +
                         " mmc.MadoguchiL1ChousaCD " +                  // 0:MagoguchiL1ChousaCD
                         ",mmc.MadoguchiL1ChousaBushoCD " +             // 1:部所CD   MadoguchiL1ChousaBushoCD
                         ",mmc.MadoguchiL1ChousaBusho " +               // 2:部所名   MadoguchiL1ChousaBusho
                         //",mb.BushokanriboKameiRaku " +                 // 3:部所略名 MadoguchiL1ChousaRyakumei
                         ",mmc.MadoguchiL1ChousaRyakumei " +            // 3:部所略名 MadoguchiL1ChousaRyakumei
                         ",mmc.MadoguchiL1ChousaTantoushaCD  " +        // 4:調査員CD MadoguchiL1ChousaTantoushaCD
                         ",mc.ChousainMei  " +                          // 5:調査員名 MadoguchiL1ChousaTantousha
                         ",FORMAT(mmc.MadoguchiL1ChousaShimekiribi,'yyyy/MM/dd') AS  Shimekiribi " +     // 6:締切日   MadoguchiL1ChousaShimekiribi
                         ",mmc.MadoguchiL1ChousaShinchoku  " +          // 7:進捗状況 MadoguchiL1ChousaShinchoku
                         ",mmc.MadoguchiL1ChousaKakunin  " +            // 8:確認     MadoguchiL1ChousaKakunin
                         ",mmc.MadoguchiL1Memo  " +                     // 9:メモ     MadoguchiL1Memo
                         ",mmc.MadoguchiL1Bunya  " +                    // 10:分野    MadoguchiL1Bunya
                         ",mmc.MadoguchiL1Bunrui  " +                   // 11:分類    MadoguchiL1Bunrui
                         ",'" + MadoguchiID + "' AS MadoguchiID " +     // 12:窓口ID  MadoguchiID

                         ",MadoguchiL1TokuchoBangou " +                 // 13:特調番号（枝番付き）
                         ",MadoguchiL1UketsukeBangou " +                // 14:特調番号
                         ",MadoguchiL1UketsukeBangouEdaban " +          // 15:枝番
                         ",MadoguchiL1ShuTantouFlag " +                 // 16:調査担当フラグ
                         ",MadoguchiL1ShuTantouHinmokusu " +            // 17:調査担当の品目数
                         ",MadoguchiL1FukuTantouFlag " +                // 18:副調査担当1,2フラグ
                         ",MadoguchiL1FukuTantouHinmokusu " +           // 19:副調査担当1,2の品目数
                         ",MadoguchiL1ShukeihyoLink " +                 // 20:集計表フォルダ

                         "FROM MadoguchiJouhouMadoguchiL1Chou mmc " +
                         "LEFT JOIN Mst_Busho mb on mb.GyoumuBushoCD = mmc.MadoguchiL1ChousaBushoCD " +
                         "LEFT JOIN Mst_Chousain mc on mc.KojinCD = mmc.MadoguchiL1ChousaTantoushaCD " +
                         "WHERE mmc.MadoguchiID = '" + MadoguchiID + "' " +
                         "order by mmc.MadoguchiID,mmc.MadoguchiL1ChousaCD ";

                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    dtMadoguchiL1 = new DataTable();
                    sda.Fill(dtMadoguchiL1);

                    // 取得データがある場合
                    if (dtMadoguchiL1.Rows.Count > 0)
                    {
                        for (int i = 0; dtMadoguchiL1.Rows.Count > i; i++)
                        {
                            // Listに詰める
                            MagoguchiL1ChousaCDList_c.Add(dtMadoguchiL1.Rows[i][0].ToString());
                            MadoguchiL1ChousaBushoCDList_c.Add(dtMadoguchiL1.Rows[i][1].ToString());
                            MadoguchiL1ChousaBushoList_c.Add(dtMadoguchiL1.Rows[i][2].ToString());
                            MadoguchiL1ChousaRyakumeiList_c.Add(dtMadoguchiL1.Rows[i][3].ToString());
                            MadoguchiL1ChousaTantoushaCDList_c.Add(dtMadoguchiL1.Rows[i][4].ToString());
                            MadoguchiL1ChousaTantoushaList_c.Add(dtMadoguchiL1.Rows[i][5].ToString());
                            MadoguchiL1ChousaShimekiribiList_c.Add(dtMadoguchiL1.Rows[i][6].ToString());
                            MadoguchiL1ChousaShinchokuList_c.Add(dtMadoguchiL1.Rows[i][7].ToString());
                            MadoguchiL1ChousaKakuninList_c.Add(dtMadoguchiL1.Rows[i][8].ToString());
                            MadoguchiL1MemoList_c.Add(dtMadoguchiL1.Rows[i][9].ToString());
                            MadoguchiL1BunyaList_c.Add(dtMadoguchiL1.Rows[i][10].ToString());
                            MadoguchiL1BunruiList_c.Add(dtMadoguchiL1.Rows[i][11].ToString());

                            MadoguchiL1TokuchoBangou_c.Add(dtMadoguchiL1.Rows[i][13].ToString());
                            MadoguchiL1UketsukeBangou_c.Add(dtMadoguchiL1.Rows[i][14].ToString());
                            MadoguchiL1UketsukeBangouEdaban_c.Add(dtMadoguchiL1.Rows[i][15].ToString());
                            MadoguchiL1ShuTantouFlag_c.Add(dtMadoguchiL1.Rows[i][16].ToString());
                            MadoguchiL1ShuTantouHinmokusu_c.Add(dtMadoguchiL1.Rows[i][17].ToString());
                            MadoguchiL1FukuTantouFlag_c.Add(dtMadoguchiL1.Rows[i][18].ToString());
                            MadoguchiL1FukuTantouHinmokusu_c.Add(dtMadoguchiL1.Rows[i][19].ToString());
                            MadoguchiL1ShukeihyoLink_c.Add(dtMadoguchiL1.Rows[i][20].ToString());

                            if (!"0".Equals(dtMadoguchiL1.Rows[i][4].ToString()))
                            {
                                i_keyList_c.Add(dtMadoguchiL1.Rows[i][1].ToString() + "," + dtMadoguchiL1.Rows[i][4].ToString()); // MadoguchiL1ChousaBushoCD + "," + MadoguchiL1ChousaTantoushaCD
                            }
                            else
                            {
                                // 担当者がいない場合（調査員CDが0の場合)、i_SDTMadoguchiL1Chousa_pと合わせる為、「MadoguchiL1ChousaBushoCD + ","」だけにする
                                i_keyList_c.Add(dtMadoguchiL1.Rows[i][1].ToString() + ","); // MadoguchiL1ChousaBushoCD + ","
                            }
                            // 2:更新 3:削除 それ以外は新規
                            i_modeList_c.Add("0");
                            i_shufukuList_c.Add("0");
                        }
                    }

                    string i_key = "";
                    num = 0;
                    // 更新・削除の対象を振り分ける
                    for (int i = 0; i_keyList_c.Count > i; i++)
                    {
                        // keyを取り出す MadoguchiL1ChousaBushoCD + "," + MadoguchiL1ChousaTantoushaCD
                        i_key = i_keyList_c[i];

                        // i_SDTMadoguchiL1Chousa_p に存在した場合、更新対象、存在しない場合、削除対象

                        if (i_SDTMadoguchiL1Chousa_p.IndexOf(i_key) > -1)
                        {
                            num = i_SDTMadoguchiL1Chousa_p.IndexOf(i_key);
                            //updmessage2 = 1;

                            // 更新
                            i_modeList_c[i] = "2";

                            i_SDTMadoguchiL1Chousa_p.RemoveAt(num);
                            MagoguchiL1ChousaCDList.RemoveAt(num);
                            MadoguchiL1ChousaBushoCDList.RemoveAt(num);
                            MadoguchiL1ChousaBushoList.RemoveAt(num);
                            MadoguchiL1ChousaRyakumeiList.RemoveAt(num);
                            MadoguchiL1ChousaTantoushaCDList.RemoveAt(num);
                            MadoguchiL1ChousaTantoushaList.RemoveAt(num);
                            MadoguchiL1ChousaShimekiribiList.RemoveAt(num);
                            MadoguchiL1ChousaShinchokuList.RemoveAt(num);

                            MadoguchiL1ShuTantouFlagList.RemoveAt(num);
                            MadoguchiL1ShuTantouHinmokusuList.RemoveAt(num);
                            MadoguchiL1FukuTantouFlagList.RemoveAt(num);
                            MadoguchiL1FukuTantouHinmokusuList.RemoveAt(num);
                            MadoguchiL1ChousaKindList.RemoveAt(num);
                        }
                        else
                        {
                            num = i_SDTMadoguchiL1Chousa_p.IndexOf(i_key);
                            updmessage3 = 1;
                            // 削除
                            i_modeList_c[i] = "3";

                            //i_SDTMadoguchiL1Chousa_p.RemoveAt(num);
                            //MagoguchiL1ChousaCDList.RemoveAt(num);
                            //MadoguchiL1ChousaBushoCDList.RemoveAt(num);
                            //MadoguchiL1ChousaBushoList.RemoveAt(num);
                            //MadoguchiL1ChousaRyakumeiList.RemoveAt(num);
                            //MadoguchiL1ChousaTantoushaCDList.RemoveAt(num);
                            //MadoguchiL1ChousaTantoushaList.RemoveAt(num);
                            //MadoguchiL1ChousaShimekiribiList.RemoveAt(num);
                            //MadoguchiL1ChousaShinchokuList.RemoveAt(num);
                        }
                    }


                    // Madoguchi：窓口ミハル Tokumei：特命課長 Jibun：自分大臣
                    string bikou = "";

                    // 更新フラグ true:更新 false:更新なし
                    Boolean updateFlg = false;
                    if ("Tokumei".Equals(gamenMode))
                    {
                        bikou = "◎";
                    }
                    else
                    {
                        bikou = "○";
                    }

                    // 調査担当者数
                    string chousaTantouCnt = "0";
                    // 調査担当者フラグ
                    string chousaTantouFlg = "0";
                    // 副調査担当者数
                    string fukuChousaTantouCnt = "0";
                    // 副調査担当者フラグ
                    string fukuChousaTantouFlg = "0";

                    // 更新処理
                    for (int i = 0; i_keyList_c.Count > i; i++)
                    {
                        // modeが2：更新のデータのみを更新する
                        if ("2".Equals(i_modeList_c[i]))
                        {
                            updateFlg = false;
                            chousaTantouCnt = "0";
                            chousaTantouFlg = "0";
                            fukuChousaTantouCnt = "0";
                            fukuChousaTantouFlg = "0";
                            // 項目に変更があるかチェック

                            // 調査員CDが空ではない場合、調査員の品目数を取得する
                            if (MadoguchiL1ChousaTantoushaCDList_c[i] != "" && MadoguchiL1ChousaTantoushaCDList_c[i] != "0")
                            {
                                cnt_DataTable = new DataTable();
                                cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " HinmokuChousainCD = '" + MadoguchiL1ChousaTantoushaCDList_c[i] + "' ");
                                if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                {
                                    chousaTantouCnt = cnt_DataTable.Rows[0][0].ToString();
                                    // カウントが0以外なら調査担当者フラグを1立てる
                                    if (!"0".Equals(chousaTantouCnt))
                                    {
                                        chousaTantouFlg = "1";
                                    }
                                }
                            }
                            // 部所のみで担当者が空の場合
                            else
                            {
                                cnt_DataTable = new DataTable();
                                cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " HinmokuChousainCD is null AND HinmokuRyakuBushoCD = '" + MadoguchiL1ChousaBushoCDList_c[i] + "' ");
                                if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                {
                                    chousaTantouCnt = cnt_DataTable.Rows[0][0].ToString();
                                    // カウントが0以外なら調査担当者フラグを1立てる
                                    if (!"0".Equals(chousaTantouCnt))
                                    {
                                        chousaTantouFlg = "1";
                                    }
                                }

                            }

                            // 調査員CDが空ではない場合、副調査員の品目数を取得する
                            if (MadoguchiL1ChousaTantoushaCDList_c[i] != "" && MadoguchiL1ChousaTantoushaCDList_c[i] != "0")
                            {
                                cnt_DataTable = new DataTable();
                                cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " ((HinmokuFukuChousainCD1 = '" + MadoguchiL1ChousaTantoushaCDList_c[i] + "' AND HinmokuRyakuBushoFuku1CD = '" + MadoguchiL1ChousaBushoCDList_c[i] + "') or (HinmokuFukuChousainCD2 = '" + MadoguchiL1ChousaTantoushaCDList_c[i] + "' AND HinmokuRyakuBushoFuku2CD = '" + MadoguchiL1ChousaBushoCDList_c[i] + "')) ");
                                if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                {
                                    fukuChousaTantouCnt = cnt_DataTable.Rows[0][0].ToString();
                                    // カウントが0以外なら副調査担当者フラグを1立てる
                                    if (!"0".Equals(fukuChousaTantouCnt))
                                    {
                                        fukuChousaTantouFlg = "1";
                                    }
                                }
                            }
                            // 部所のみで担当者が空の場合
                            else
                            {
                                cnt_DataTable = new DataTable();
                                //cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " (HinmokuFukuChousainCD1 is null or HinmokuFukuChousainCD2 is null) ");
                                cnt_DataTable = getData(cnt_discript, cnt_value, cnt_table, cnt_where + " ((HinmokuFukuChousainCD1 is null AND HinmokuRyakuBushoFuku1CD = '" + MadoguchiL1ChousaBushoCDList_c[i] + "') or (HinmokuFukuChousainCD2 is null AND HinmokuRyakuBushoFuku2CD = '" + MadoguchiL1ChousaBushoCDList_c[i] + "')) ");
                                if (cnt_DataTable != null && cnt_DataTable.Rows.Count > 0)
                                {
                                    fukuChousaTantouCnt = cnt_DataTable.Rows[0][0].ToString();
                                    // カウントが0以外なら副調査担当者フラグを1立てる
                                    if (!"0".Equals(fukuChousaTantouCnt))
                                    {
                                        fukuChousaTantouFlg = "1";
                                    }
                                }
                            }

                            // 差異があるかチェック
                            if (!chousaTantouFlg.Equals(MadoguchiL1ShuTantouFlag_c[i]) || !chousaTantouCnt.Equals(MadoguchiL1ShuTantouHinmokusu_c[i])
                                || !fukuChousaTantouFlg.Equals(MadoguchiL1FukuTantouFlag_c[i]) || !fukuChousaTantouCnt.Equals(MadoguchiL1FukuTantouHinmokusu_c[i]))
                            {
                                updateFlg = true;
                            }

                            // 項目に変更がある場合にのみ更新を行う
                            if (updateFlg == true)
                            {

                                updmessage2 = 1;
                                cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                                    " MadoguchiL1ChousaBushoCD = '" + MadoguchiL1ChousaBushoCDList_c[i] + "' " +
                                    ",MadoguchiL1ChousaBusho = mb.BushoKanriboKamei " +
                                    ",MadoguchiL1ChousaBushoOld = N'" + MadoguchiL1ChousaBushoList_c[i] + "' " +
                                    ",MadoguchiL1ChousaRyakumei = mb.BushoKanriboKameiRaku " +
                                    ",MadoguchiL1ChousaRyakumeiOld = N'" + MadoguchiL1ChousaRyakumeiList_c[i] + "' " +
                                    ",MadoguchiL1ChousaTantoushaCD = '" + MadoguchiL1ChousaTantoushaCDList_c[i] + "' " +
                                    ",MadoguchiL1ChousaTantousha = mc.ChousainMei " +
                                    ",MadoguchiL1ChousaTantoushaOld = N'" + MadoguchiL1ChousaTantoushaList_c[i] + "' " +
                                    ",MadoguchiL1ChousaShimekiribi = '" + MadoguchiL1ChousaShimekiribiList_c[i] + "' " +

                                    ",MadoguchiL1ShuTantouFlag = " + chousaTantouFlg + "" +
                                    ",MadoguchiL1ShuTantouHinmokusu = " + chousaTantouCnt + "" +
                                    ",MadoguchiL1FukuTantouFlag = " + fukuChousaTantouFlg + "" +
                                    ",MadoguchiL1FukuTantouHinmokusu = " + fukuChousaTantouCnt + "" +

                                    ",MadoguchiL1AsteriaKoushinFlag = 1 " +  // Asteria更新フラグ 品目本数等の値が変わったら1にする
                                    ",MadoguchiL1TokuchoHaitaFlag = 1 " +  // 898 対応 品目数が変わったら排他フラグも1:ONにする

                                    //No1719 ShiryouHolderも更新する
                                    ",MadoguchiL1SiryouHolder = '" + ShiryouHolder + "' " +

                                    ",MadoguchiL1UpdateDate = SYSDATETIME() " +
                                    ",MadoguchiL1UpdateUser = N'" + UpdateKojinCD + "' " +
                                    ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
                                    "FROM MadoguchiJouhouMadoguchiL1Chou mjm " +
                                    "LEFT JOIN Mst_Busho mb ON mb.GyoumuBushoCD = mjm.MadoguchiL1ChousaBushoCD " +
                                    "LEFT JOIN Mst_Chousain mc ON mc.KojinCD = mjm.MadoguchiL1ChousaTantoushaCD " +
                                    "WHERE MadoguchiID = '" + MadoguchiID + "' " +
                                    "AND MadoguchiL1ChousaBushoCD = '" + MadoguchiL1ChousaBushoCDList_c[i] + "' " +
                                    "AND MadoguchiL1ChousaCD = '" + MagoguchiL1ChousaCDList_c[i] + "' ";
                                cmd.ExecuteNonQuery();

                                // 支部備考の更新
                                if (MadoguchiL1ChousaBushoCDList_c[i] != "")
                                {
                                    // 支部備考更新
                                    ProUpdateBikoFromBusho(MadoguchiID, gamenMode, UpdateKojinCD, bikou, MadoguchiL1ChousaBushoCDList_c[i]);
                                }
                            }
                        }
                        // modeが3：削除のデータを削除する
                        else if ("3".Equals(i_modeList_c[i]))
                        {
                            cmd.CommandText = "DELETE FROM MadoguchiJouhouMadoguchiL1Chou " +
                                "WHERE MadoguchiID = '" + MadoguchiID + "' " +
                                "AND MadoguchiL1ChousaBushoCD = '" + MadoguchiL1ChousaBushoCDList_c[i] + "' " +
                                "AND MadoguchiL1ChousaCD = '" + MagoguchiL1ChousaCDList_c[i] + "' ";
                            cmd.ExecuteNonQuery();

                            // 支部備考の更新
                            if (MadoguchiL1ChousaBushoCDList_c[i] != "")
                            {
                                // 支部備考更新
                                ProUpdateBikoFromBusho(MadoguchiID, gamenMode, UpdateKojinCD, bikou, MadoguchiL1ChousaBushoCDList_c[i]);
                            }
                        }
                    }

                    // MadoguchiL1ChousaCD の最大値を取得する
                    string madoguchiL1TantouIDMaxCD = "";
                    int maxCD = 0;
                    int shinchoku = 10; // 0が無くなって10:依頼スタート

                    cmd.CommandText = "SELECT  " +
                     "Max(MadoguchiL1ChousaCD) AS MaxCD " +     // 0:MadoguchiL1ChousaCD
                     "FROM MadoguchiJouhouMadoguchiL1Chou " +
                     "WHERE MadoguchiID = '" + MadoguchiID + "' ";

                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    dtMadoguchiL1 = new DataTable();
                    sda.Fill(dtMadoguchiL1);

                    if (dtMadoguchiL1.Rows.Count > 0)
                    {
                        madoguchiL1TantouIDMaxCD = dtMadoguchiL1.Rows[0][0].ToString();
                        // 数値変換
                        int.TryParse(madoguchiL1TantouIDMaxCD, out maxCD);
                    }

                    int index_row = 0;
                    // 新規登録（i_SDTMadoguchiL1Chousa_pで取得したデータを基にする）
                    for (int i = 0; i_SDTMadoguchiL1Chousa_p.Count > i; i++)
                    {
                        updmessage1 = 1;
                        maxCD += 1;

                        // 進捗を判断
                        if (!"".Equals(MadoguchiL1ChousaBushoCDList[i]) && "".Equals(MadoguchiL1ChousaTantoushaCDList[i]))
                        {
                            // 0:依頼 ⇒ 10:依頼
                            shinchoku = 10;
                        }
                        else if (!"".Equals(MadoguchiL1ChousaBushoCDList[i]) && !"".Equals(MadoguchiL1ChousaTantoushaCDList[i]))
                        {
                            // 1:調査中 ⇒ 40:集計中
                            //shinchoku = 40;
                            // 30:見積中
                            // 20:調査開始
                            shinchoku = 20;
                        }

                        cmd.CommandText = "INSERT INTO MadoguchiJouhouMadoguchiL1Chou (" +
                            " MadoguchiID " +
                            ",MadoguchiL1ChousaCD " +
                            ",MadoguchiL1ChousaBushoCD " +
                            ",MadoguchiL1ChousaBusho " +
                            ",MadoguchiL1ChousaBushoOld " +
                            ",MadoguchiL1ChousaTantoushaCD " +
                            ",MadoguchiL1ChousaTantousha " +
                            ",MadoguchiL1ChousaTantoushaOld " +
                            ",MadoguchiL1ChousaShimekiribi " +
                            ",MadoguchiL1ChousaShinchoku " +
                            ",MadoguchiL1ChousaKakunin " +
                            ",MadoguchiL1Memo " +
                            ",MadoguchiL1Bunya " +
                            ",MadoguchiL1Bunrui " +
                            ",MadoguchiL1KanrihyoHaitaFlag " +
                            ",MadoguchiL1ShinchokuHaitaFlag " +
                            ",MadoguchiL1TokuchoHaitaFlag " +
                            ",MadoguchiL1AsteriaKoushinFlag " +
                            ",MadoguchiL1TokuchoBangou " +
                            ",MadoguchiL1UketsukeBangou " +
                            ",MadoguchiL1UketsukeBangouEdaban " +
                            ",MadoguchiL1ShuTantouFlag " +
                            ",MadoguchiL1ShuTantouHinmokusu " +
                            ",MadoguchiL1FukuTantouFlag " +
                            ",MadoguchiL1FukuTantouHinmokusu " +
                            ",MadoguchiL1MitsumoriFrom " +
                            //",MadoguchiL1MitsumoriTo " +
                            ",MadoguchiL1ShukeihyoLink " +
                            ",MadoguchiL1TourokuNendo " +
                            ",MadoguchiL1CreateDate " +
                            ",MadoguchiL1CreateUser " +
                            ",MadoguchiL1CreateProgram " +
                            ",MadoguchiL1UpdateDate " +
                            ",MadoguchiL1UpdateUser " +
                            ",MadoguchiL1UpdateProgram " +
                            ",MadoguchiL1DeleteFlag " +
                            ",MadoguchiL1HachushaMei " +
                            ",MadoguchiL1ChousaRyakumei " +
                            ",MadoguchiL1ChousaRyakumeiOld " +
                            // No1719 MadoguchiJouhouMadoguchiL1Chouの[MadoguchiL1SiryouHolder]が特調奉行、工程まもる側に登録されない。
                            ",MadoguchiL1SiryouHolder " +
                            ") VALUES ( " +
                            "'" + MadoguchiID + "' " +                              // MadoguchiID
                            ",'" + maxCD + "' " +                                   // MadoguchiL1ChousaCD
                            ",'" + MadoguchiL1ChousaBushoCDList[i] + "' " +         // MadoguchiL1ChousaBushoCD
                            ",N'" + MadoguchiL1ChousaBushoList[i] + "' " +           // MadoguchiL1ChousaBusho
                            ",N'" + MadoguchiL1ChousaBushoList[i] + "' " +           // MadoguchiL1ChousaBushoOld
                            ",'" + MadoguchiL1ChousaTantoushaCDList[i] + "' " +     // MadoguchiL1ChousaTantoushaCD
                            ",N'" + MadoguchiL1ChousaTantoushaList[i] + "' " +       // MadoguchiL1ChousaTantousha
                            ",N'" + MadoguchiL1ChousaTantoushaList[i] + "' " +       // MadoguchiL1ChousaTantoushaOld
                            ",'" + MadoguchiL1ChousaShimekiribiList[i] + "' " +     // MadoguchiL1ChousaShimekiribi
                            ",'" + shinchoku + "' " +                               // MadoguchiL1ChousaShinchoku
                            ",'0' " +                                               // MadoguchiL1ChousaKakunin
                            ",'' " +                                                // MadoguchiL1Memo
                            ",'' " +                                                // MadoguchiL1Bunya
                            ",'' " +                                                // MadoguchiL1Bunrui
                            ",'0' " +                                               // MadoguchiL1KanrihyoHaitaFlag
                            ",'0' " +                                               // MadoguchiL1ShinchokuHaitaFlag
                            ",'0' " +                                               // MadoguchiL1TokuchoHaitaFlag
                            ",'1' " +                                               // MadoguchiL1AsteriaKoushinFlag
                            ",N'" + TokuchoBangou + "' " +                           // MadoguchiL1TokuchoBangou
                            ",N'" + UketsukeBangou + "' " +                          // MadoguchiL1UketsukeBangou
                            ",N'" + ChangeSqlText(UketsukeBangouEdaban, 1) + "' " +  // MadoguchiL1UketsukeBangouEdaban・・・枝番は自由入力なのでエスケープ
                            ",'" + MadoguchiL1ShuTantouFlagList[i] + "' " +         // MadoguchiL1ShuTantouFlag
                            ",'" + MadoguchiL1ShuTantouHinmokusuList[i] + "' " +    // MadoguchiL1ShuTantouHinmokusu
                            ",'" + MadoguchiL1FukuTantouFlagList[i] + "' " +        // MadoguchiL1FukuTantouFlag
                            ",'" + MadoguchiL1FukuTantouHinmokusuList[i] + "' " +   // MadoguchiL1FukuTantouHinmokusu
                            //",'" + DateTime.Today + "' " +                          // MadoguchiL1MitsumoriFrom
                            ",'" + Tourokubi + "' " +                               // MadoguchiL1MitsumoriFrom
                            //",'" + Shimekiribi + "' " +                             // MadoguchiL1MitsumoriTo
                            ",N'" + ShukeihyoLink + "' " +                           // MadoguchiL1ShukeihyoLink
                            ",'" + TourokuNendo + "' " +                            // MadoguchiL1TourokuNendo
                            ",SYSDATETIME() " +                                     // 登録日時
                            ",N'" + UpdateKojinCD + "' " +                           // 登録ユーザ
                            ",'" + pgmName + methodName + "' " +                    // 登録プログラム
                            ",SYSDATETIME() " +                                     // 更新日時
                            ",N'" + UpdateKojinCD + "' " +                           // 更新ユーザ
                            ",'" + pgmName + methodName + "' " +                    // 更新プログラム
                            ",0 " +                                                 // 削除フラグ
                            ",N'" + HachuuKikanmei + "'" +                           // MadoguchiHachuuKikanmei
                            ",N'" + MadoguchiL1ChousaRyakumeiList[i] + "' " +        // MadoguchiL1ChousaRyakumei
                            ",N'" + MadoguchiL1ChousaRyakumeiList[i] + "' " +        // MadoguchiL1ChousaRyakumei
                            //No1719 MadoguchiJouhouMadoguchiL1Chouの[MadoguchiL1SiryouHolder]が特調奉行、工程まもる側に登録されない。
                            ",N'" + ShiryouHolder + "' " +                          // MadoguchiL1SiryouHolder
                            ") ";
                        cmd.ExecuteNonQuery();

                        // ログ出力
                        outputLogger("MadoguchiHinmokuRenkeiUpdate_SQL", "New Max = " + maxCD, "debug", UpdateKojinCD);

                        // 支部備考の更新
                        ProUpdateBikoFromBusho(MadoguchiID, gamenMode, UpdateKojinCD, bikou, MadoguchiL1ChousaBushoCDList[i]);

                        index_row = i;
                    }

                    // ここでコミットしておかないと後続でselect出来ない
                    transaction.Commit();
                    transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    // 担当部所毎に行うので、ここで削除対象を更新する
                    string i_deleteCheck = "0";
                    dtMadoguchiL1 = new DataTable();
                    if (MadoguchiL1ChousaBushoCDList.Count > 0) 
                    { 
                        cmd.CommandText = "SELECT  " +
                         "MadoguchiID " +     // 0:MadoguchiID
                         "FROM MadoguchiJouhouMadoguchiL1Chou " +
                         "WHERE MadoguchiID = '" + MadoguchiID + "' " +
                         "AND MadoguchiL1ChousaBushoCD = '" + MadoguchiL1ChousaBushoCDList[index_row] + "' " +
                         "order by MadoguchiL1ChousaCD ";

                        //データ取得
                        sda = new SqlDataAdapter(cmd);
                        //dtMadoguchiL1 = new DataTable();
                        sda.Fill(dtMadoguchiL1);
                    }

                    if (dtMadoguchiL1 != null && dtMadoguchiL1.Rows.Count > 0)
                    {
                        i_deleteCheck = "1";
                    }

                    if ("1".Equals(i_deleteCheck))
                    {
                        // 支部備考の更新
                        ProUpdateBikoFromBusho(MadoguchiID, gamenMode, UpdateKojinCD, bikou, MadoguchiL1ChousaBushoCDList[index_row]);
                    }
                    
                    // 担当部所に追加された調査部所、副部所1、副部所2で支部応援をみて、GaroonTsuikaAteskiにユーザーを登録する
                    cmd.CommandText = "SELECT DISTINCT " +
                    " mc.KojinCD " +
                    ",mc.ChousainMei " +
                    ",mjmc.MadoguchiL1ChousaBushoCD " +
                    ",mb.BushokanriboKamei " +
                    "FROM MadoguchiJouhouMadoguchiL1Chou mjmc " +
                    "INNER JOIN Mst_Busho mb ON mb.GyoumuBushoCD = mjmc.MadoguchiL1ChousaBushoCD AND ISNULL(BushoDeleteFlag,0) = 0 " +
                    "INNER JOIN Mst_Chousain mc ON mc.GyoumuBushoCD = mb.GyoumuBushoCD AND ISNULL(ChousainDeleteFlag,0) = 0 " +
                    //不具合No1356(1124)　支部応援の削除フラグを見ていないところを復活。
                    "INNER JOIN Mst_Shibuouen ms ON ms.ShibuouenKojinCD = mc.KojinCD AND ISNULL(ShibuouenDeleteFlag,0) = 0 " +
                    //"INNER JOIN Mst_Shibuouen ms ON ms.ShibuouenKojinCD = mc.KojinCD " + // 支部応援の削除フラグは見ない・・・
                    "WHERE mjmc.MadoguchiID = '" + MadoguchiID + "' " +
                    "ORDER BY mjmc.MadoguchiL1ChousaBushoCD,mc.KojinCD";

                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    DataTable garoonDt = new DataTable();
                    sda.Fill(garoonDt);

                    // GaroonTuikaAtesakiを取得しておく
                    var tmpdt = new DataTable();
                    tmpdt = getData("GaroonTsuikaAtesakiBushoCD", "GaroonTsuikaAtesakiTantoushaCD", "GaroonTsuikaAtesaki", "GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "' AND GaroonTsuikaAtesakiDeleteFlag = 0 ");

                    string KojinCD = "";
                    string ChousainMei = "";
                    string BushoCD = "";
                    string BushoMei = "";
                    string where = "";
                    // 宛先追加フラグ
                    Boolean insertFlg = true;

                    if (garoonDt != null && garoonDt.Rows.Count > 0)
                    {
                        for (int i = 0; i < garoonDt.Rows.Count; i++)
                        {
                            KojinCD = garoonDt.Rows[i][0].ToString();
                            ChousainMei = garoonDt.Rows[i][1].ToString();
                            BushoCD = garoonDt.Rows[i][2].ToString();
                            BushoMei = garoonDt.Rows[i][3].ToString();

                            insertFlg = true;

                            // 既にいるか確認
                            if (tmpdt != null && tmpdt.Rows.Count > 0)
                            {
                                for(int j = 0; j < tmpdt.Rows.Count; j++)
                                {
                                    // 0:GaroonTsuikaAtesakiTantoushaCD
                                    // 1:GaroonTsuikaAtesakiBushoCD
                                    if (tmpdt.Rows[j][0].ToString() == KojinCD && tmpdt.Rows[j][1].ToString() == BushoCD)
                                    {
                                        insertFlg = false;
                                        break;
                                    }
                                }
                            }

                            // 存在しない場合、登録
                            if (insertFlg == true)
                            {
                                // GaroonTsuikaAtesakiに登録
                                cmd.CommandText = "INSERT INTO GaroonTsuikaAtesaki ( " +
                                " GaroonTsuikaAtesakiID " +
                                ",GaroonTsuikaAtesakiMadoguchiID " +
                                ",GaroonTsuikaAtesakiBushoCD " +
                                ",GaroonTsuikaAtesakiBusho " +
                                ",GaroonTsuikaAtesakiTantoushaCD " +
                                ",GaroonTsuikaAtesakiTantousha " +
                                ",GaroonTsuikaAtesakiCreateDate " +
                                ",GaroonTsuikaAtesakiCreateUser " +
                                ",GaroonTsuikaAtesakiCreateProgram " +
                                ",GaroonTsuikaAtesakiUpdateDate " +
                                ",GaroonTsuikaAtesakiUpdateUser " +
                                ",GaroonTsuikaAtesakiUpdateProgram " +
                                ",GaroonTsuikaAtesakiDeleteFlag " +
                                //不具合No1332(1084)
                                ",GaroonTsuikaAtesakiGamenFlag " +
                                ") VALUES (" +
                                "'" + getSaiban("GaroonTsuikaAtesakiID") + "' " + // GaroonTsuikaAtesakiID
                                ",'" + MadoguchiID + "' " +      　　　    // GaroonTsuikaAtesakiMadoguchiID
                                ",'" + BushoCD + "' " +          　　　    // GaroonTsuikaAtesakiBushoCD
                                ",N'" + BushoMei + "' " +                   // GaroonTsuikaAtesakiBusho
                                ",'" + KojinCD + "' " +                    // GaroonTsuikaAtesakiTantoushaCD
                                ",N'" + ChousainMei + "' " +                // GaroonTsuikaAtesakiTantousha
                                ",SYSDATETIME() " +                        // 登録日時
                                ",N'" + UpdateKojinCD + "' " +              // 登録ユーザ
                                ",'" + pgmName + methodName + "' " +       // 登録プログラム
                                ",SYSDATETIME() " +                        // 更新日時
                                ",N'" + UpdateKojinCD + "' " +              // 更新ユーザ
                                ",'" + pgmName + methodName + "' " +       // 更新プログラム
                                ",0 " +                                    // 削除フラグ
                                //不具合No1332(1084)
                                ",0 " +                                    //画面登録フラグ。ここは画面じゃないので0
                                ") ";

                                cmd.ExecuteNonQuery();
                            }
                        }
                    }

                    // 1016 部所が削除されても備考が削除されない
                    cmd.CommandText = "DELETE FROM ShibuBikou Where MadoguchiID = '" + MadoguchiID + "' " +
                                      "AND ShibuBikouBushoKanriboBushoCD not in (SELECT DISTINCT MadoguchiL1ChousaBushoCD FROM MadoguchiJouhouMadoguchiL1Chou WHERE MadoguchiID = '" + MadoguchiID + "') ";
                    cmd.ExecuteNonQuery();

                    //不具合No1332(1084)
                    // 部所が削除されてもGaroon連携の担当が削除されない　ここは物理削除のはず。更新時に画面から消えたら物理削除しているので。
                    //cmd.CommandText = "DELETE FROM GaroonTsuikaAtesaki Where GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "' " +
                    //                  "AND GaroonTsuikaAtesakiGamenFlag = 0 " +
                    //                  "AND GaroonTsuikaAtesakiBushoCD not in (SELECT DISTINCT MadoguchiL1ChousaBushoCD FROM MadoguchiJouhouMadoguchiL1Chou WHERE MadoguchiID = '" + MadoguchiID + "') ";
                    //不具合No1448(1237)
                    //窓口部署はGaroon連携から外さない対応
                    cmd.CommandText = "DELETE FROM GaroonTsuikaAtesaki Where GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "' " +
                                      "AND GaroonTsuikaAtesakiGamenFlag = 0 " +
                                      "AND GaroonTsuikaAtesakiBushoCD not in (SELECT DISTINCT MadoguchiL1ChousaBushoCD FROM MadoguchiJouhouMadoguchiL1Chou WHERE MadoguchiID = '" + MadoguchiID + "') " +
                                      "AND GaroonTsuikaAtesakiBushoCD not in (SELECT DISTINCT MadoguchiTantoushaBushoCD FROM MadoguchiJouhou WHERE MadoguchiID = '" + MadoguchiID + "') ";
                    //cmd.CommandText = "UPDATE GaroonTsuikaAtesaki SET GaroonTsuikaAtesakiDeleteFlag=1 Where GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "' " +
                    //                  "AND GaroonTsuikaAtesakiBushoCD not in (SELECT DISTINCT MadoguchiL1ChousaBushoCD FROM MadoguchiJouhouMadoguchiL1Chou WHERE MadoguchiID = '" + MadoguchiID + "') ";
                    cmd.ExecuteNonQuery();

                    // 683
                    // 集計表を全て価格を設定し、単価取込で取り込んだが担当部所の進捗状況が、担当者済みにならない。
                    // →現行にない処理なので、以下の流れで処理を実装する
                    //   １．調査担当者を重複なしで取得
                    //   ２．進捗の最小値を取得
                    //   ３．担当部所の進捗を取得し、比較
                    //   ４．差異がある場合、調査品目の最小値で担当部所の進捗を更新
                    //       差異が無い場合、スルー
                    //   
                    // 調査担当者の進捗が更新された場合、担当部所テーブルのMadoguchiL1AsteriaKoushinFlagを1にする。

                    // １．調査担当者を重複なしで取得　２．進捗の最小値を取得
                    cmd.CommandText = "SELECT DISTINCT " +
                    " MadoguchiID,HinmokuChousainCD,min(ChousaShinchokuJoukyou) AS minShinchock " +
                    "FROM ChousaHinmoku " +
                    "WHERE MadoguchiID = '" + MadoguchiID + "' AND HinmokuChousainCD is not null " +
                    "GROUP BY MadoguchiID,HinmokuChousainCD " +
                    "ORDER BY HinmokuChousainCD ";
                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    DataTable chousaTantoushaDt = new DataTable();
                    sda.Fill(chousaTantoushaDt);

                    if(chousaTantoushaDt != null && chousaTantoushaDt.Rows.Count > 0)
                    {
                        for (int i = 0; i < chousaTantoushaDt.Rows.Count; i++) { 
                            // ３．担当部所の進捗を取得し、比較
                            cmd.CommandText = "SELECT TOP 1" +
                                " MadoguchiL1ChousaShinchoku " +
                                "FROM MadoguchiJouhouMadoguchiL1Chou " +
                                "WHERE MadoguchiID = '" + MadoguchiID + "' AND MadoguchiL1ChousaTantoushaCD = '" + chousaTantoushaDt.Rows[i][1].ToString() + "' ";

                            //データ取得
                            sda = new SqlDataAdapter(cmd);
                            // DataTableはループの中で生成する
                            DataTable L1ShinchockDt = new DataTable();
                            sda.Fill(L1ShinchockDt);

                            if (L1ShinchockDt != null && L1ShinchockDt.Rows.Count > 0)
                            {
                                // 担当部所と調査品目の進捗を比較
                                if(L1ShinchockDt.Rows[0][0].ToString() != chousaTantoushaDt.Rows[i][2].ToString())
                                {
                                    int L1Shinchock = 0;
                                    int hinmokuShinchock = 0;
                                    int.TryParse(L1ShinchockDt.Rows[0][0].ToString(), out L1Shinchock);
                                    int.TryParse(chousaTantoushaDt.Rows[i][2].ToString(), out hinmokuShinchock);

                                    // 1194 担当部所の進捗よりも、調査品目の進捗が大きい場合のみ更新
                                    if(L1Shinchock <= hinmokuShinchock)
                                    {
                                        // 担当部所の進捗状況を調査品目の最小の進捗で更新
                                        cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                                            " MadoguchiL1ChousaShinchoku = '" + chousaTantoushaDt.Rows[i][2].ToString() + "' " +
                                            ",MadoguchiL1AsteriaKoushinFlag = 1" + // Asteria更新フラグを更新
                                            ",MadoguchiL1UpdateDate = SYSDATETIME() " +
                                            ",MadoguchiL1UpdateUser = N'" + UpdateKojinCD + "' " +
                                            ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
                                            "WHERE MadoguchiID = '" + MadoguchiID + "' AND MadoguchiL1ChousaTantoushaCD = '" + chousaTantoushaDt.Rows[i][1].ToString() + "' ";
                                        cmd.ExecuteNonQuery();

                                        updmessage2 = 1;
                                    }
                                }
                            }
                        }
                    }

                    // 担当部所内が更新されたときの窓口情報の連動更新
                    cmd.CommandText = "SELECT  " +
                         "Min(MadoguchiL1ChousaShinchoku) AS MinShinchoku " +     // 0:MadoguchiL1ChousaShinchoku
                         "FROM MadoguchiJouhouMadoguchiL1Chou " +
                         "WHERE MadoguchiID = '" + MadoguchiID + "' ";

                    //データ取得
                    sda = new SqlDataAdapter(cmd);
                    dtMadoguchiL1 = new DataTable();
                    sda.Fill(dtMadoguchiL1);

                    if (dtMadoguchiL1.Rows.Count > 0)
                    {
                        // 最小の進捗取り出し
                        int.TryParse(dtMadoguchiL1.Rows[0][0].ToString(), out shinchoku);
                    }

                    // MadoguchiJouhou の進捗状況を更新
                    cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                        " MadoguchiShinchokuJoukyou = '" + shinchoku + "' " +
                        ",MadoguchiUpdateDate = SYSDATETIME()" +
                        ",MadoguchiUpdateUser = N'" + UpdateKojinCD + "' " +
                        ",MadoguchiUpdateProgram = '" + pgmName + methodName + "' " +
                        "WHERE MadoguchiID = '" + MadoguchiID + "' " +
                        "AND MadoguchiShinchokuJoukyou <= '" + shinchoku + "' ";
                    cmd.ExecuteNonQuery();

                    if (updmessage1 == 1)
                    {
                        // I20304:担当部所へ新規行追加があります。
                        mes += GetMessage("I20304", "") + Environment.NewLine;
                    }
                    // 現行で更新、削除でもI20306:「担当部所の行が削除されました。」を出していたのを分離
                    if (updmessage2 == 1)
                    {
                        // I20305:担当部所へ更新があります。
                        mes += GetMessage("I20305", "") + Environment.NewLine;
                    }
                    if (updmessage3 == 1)
                    {
                        // I20306:担当部所の行が削除されました。
                        mes += GetMessage("I20306", "") + Environment.NewLine;
                    }

                }
                catch (ArithmeticException e)
                {
                    transaction.Rollback();
                    Console.WriteLine(e);
                    return false;
                }
                catch (Exception e)
                {
                    return false;
                }
                finally
                {
                    transaction.Commit();
                    conn.Close();
                }

                DataTable dt = new DataTable();
                dt = getData("GyoumuBushoCD", "GyoumuBushoCD", "Mst_Chousain", "KojinCD = " + UpdateKojinCD);
                string GyoumuBushoCD = "";
                if (dt != null && dt.Rows.Count > 0)
                {
                    GyoumuBushoCD = dt.Rows[0][0].ToString();
                }

                // 皇帝まもる連携
                KouteiTantouBushoRenkei(MadoguchiID, UpdateKojinCD, GyoumuBushoCD);
            }
            return true;
        }

        // 支部備考更新
        private void ProUpdateBikoFromBusho(string MadoguchiID, string gamenMode, string UpdateKojinCD, String bikou, String MadoguchiL1ChousaBushoCD)
        {
            string methodName = ".ProUpdateBikoFromBusho";

            using (var conn2 = new SqlConnection(connStr))
            {
                conn2.Open();
                var cmd = conn2.CreateCommand();

                // データがあるかないか確認の為、SELECT
                cmd.CommandText = "SELECT  " +
                     "ShibuBikouID " +                   // 0:ShibuBikouID
                     "FROM ShibuBikou " +
                     "WHERE MadoguchiID = '" + MadoguchiID + "' " +
                     "AND ShibuBikouBushoKanriboBushoCD = '" + MadoguchiL1ChousaBushoCD + "' " +
                     "order by ShibuBikouID ";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                DataTable dtMadoguchiL1 = new DataTable();
                sda.Fill(dtMadoguchiL1);

                if (dtMadoguchiL1.Rows.Count > 0)
                {
                    // 存在するので、更新
                    cmd.CommandText = "UPDATE ShibuBikou SET " +
                        " ShibuBikouChousaBusho = N'" + bikou + "' " +
                        ",ShibuBikouUpdateDate = SYSDATETIME() " +
                        ",ShibuBikouUpdateUser = N'" + UpdateKojinCD + "' " +
                        ",ShibuBikouUpdateProgram = '" + pgmName + methodName + "' " +
                        "WHERE MadoguchiID = '" + MadoguchiID + "' " +
                        "AND ShibuBikouBushoKanriboBushoCD = '" + MadoguchiL1ChousaBushoCD + "' " +
                    cmd.ExecuteNonQuery();
                }
                else
                {
                    // 存在しない為、新規登録
                    //getSaiban("ShibuBikouID")
                    cmd.CommandText = "INSERT INTO ShibuBikou(" +
                        " MadoguchiID " +
                        ",ShibuBikouID " +
                        ",ShibuBikouBushoKanriboBushoCD " +
                        ",ShibuBikouKanriNo " +
                        ",ShibuBikouChousaBusho " +
                        ",ShibuBikou " +
                        ",ShibuBikouRyakumei " +
                        ",ShibuBikouCreateDate " +
                        ",ShibuBikouCreateProgram " +
                        ",ShibuBikouCreateUser " +
                        ",ShibuBikouUpdateDate " +
                        ",ShibuBikouUpdateProgram " +
                        ",ShibuBikouUpdateUser " +
                        ",ShinDeleteFlag " +
                        ",ShibuBushokanriboKameiRakuOld " +
                        ",ShibuBushoKanriboShibuMeiOld " +
                        ") " +
                        " VALUES (" +
                        " '" + MadoguchiID + "' " +                         // MadoguchiID 
                        ",'" + getSaiban("ShibuBikouID") + "' " +           // ShibuBikouID 
                        ",'" + MadoguchiL1ChousaBushoCD + "' " +            // ShibuBikouBushoKanriboBushoCD 
                        ",null " +                                            // ShibuBikouKanriNo 
                        ",N'" + bikou + "' " +                               // ShibuBikouChousaBusho 
                        ",null ";                                             // ShibuBikou 

                    if(MadoguchiL1ChousaBushoCD != "") 
                    { 
                        cmd.CommandText += ",(SELECT BushokanriboKameiRaku FROM MST_Busho WHERE GyoumuBushoCD = '" + MadoguchiL1ChousaBushoCD + "') ";  // ShibuBikouRyakumei
                    }
                    else
                    {
                        cmd.CommandText += ",null ";  // ShibuBikouRyakumei
                    }

                    cmd.CommandText += ",SYSDATETIME() " +                  // ShibuBikouCreateDate 登録日時
                        ",'" + pgmName + methodName + "' " +                // ShibuBikouCreateProgram 登録プログラム
                        ",N'" + UpdateKojinCD + "' " +                       // ShibuBikouCreateUser 登録ユーザ
                        ",SYSDATETIME() " +                                 // ShibuBikouUpdateDate 更新日時
                        ",'" + pgmName + methodName + "' " +                // ShibuBikouUpdateProgram 更新プログラム
                        ",N'" + UpdateKojinCD + "' " +                       // ShibuBikouUpdateUser 更新ユーザ
                        ",0 " +                                             // 削除フラグ
                        ",'' " +                                            // ShibuBushokanriboKameiRakuOld 
                        ",'' " +                                            // ShibuBushoKanriboShibuMeiOld
                        ") ";
                    cmd.ExecuteNonQuery();
                }
                conn2.Close();
            }
        }



        //現在年度の取得
        public string GetTodayNendo()
        {
            string Nendo = DateTime.Today.Year.ToString();

            String discript = "NendoSeireki";
            String value = "NendoID ";
            String table = "Mst_Nendo";
            // CONVERT ( VARCHAR, GETDATE(), 111 )　・・・現在日時を年月日のフォーマットに変換する
            String where = "Nendo_Sdate <= CONVERT ( VARCHAR, GETDATE(), 111 ) AND Nendo_EDate >= CONVERT ( VARCHAR, GETDATE(), 111 )";
            DataTable dt = getData(discript, value, table, where);
            if (dt != null && dt.Rows.Count > 0)
            {
                Nendo = dt.Rows[0][0].ToString();
            }

            return Nendo;
        }

        private decimal GetLong(string str)
        {
            decimal num = 0;
            decimal.TryParse(str.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty), out num);
            return num;
        }

        // 印刷履歴の出力
        public void Insert_PrintHistory(string[] data)
        {
            SqlConnection sqlconn = new SqlConnection(connStr);
            sqlconn.Open();
            var cmd = sqlconn.CreateCommand();
            SqlTransaction transaction = sqlconn.BeginTransaction();
            cmd.Transaction = transaction;

            try
            {
                int PrintHistoryID = getSaiban("PrintHistoryID");

                cmd.CommandText = "INSERT INTO T_PrintHistory("
                                + " PrintHistoryID"                     // 1.印刷履歴ID
                                + ", PrintHistoryDateTime"              // 2.印刷日時
                                + ", PrintBushoCD"                      // 3.部所CD
                                + ", PrintBushoName"                    // 4.部所名
                                + ", PrintKojinCD"                      // 5.個人CD
                                + ", PrintUserName"                     // 6.職員名
                                + ", PrintBusinessName"                 // 7.業務名
                                + ", PrintTokuchouName"                 // 8.特調奉行名
                                + ", PrintFunctionName"                 // 9.画面・機能名
                                + ", PrintPatternName"                  // 10.帳票分類名
                                + ", PrintHistoryName"                  // 11.帳票名
                                + ", PrintHistoryFileName"              // 12.雛型ファイル名
                                + ", PrintHistoryDownLoadFileName"      // 13.ダウンロードファイル名
                                + ") VALUES ("
                                + " " + PrintHistoryID                  // 1.印刷履歴ID
                                + ", " + "SYSDATETIME()"                // 2.印刷日時
                                + ", " + "'" + data[0] + "'"            // 3.部所CD
                                + ", " + "N'" + data[1] + "'"            // 4.部所名
                                + ", " + "'" + data[2] + "'"            // 5.個人CD
                                + ", " + "N'" + data[3] + "'"            // 6.職員名
                                + ", " + "N'" + data[4] + "'"            // 7.業務名
                                + ", " + "N'" + data[5] + "'"            // 8.特調奉行名
                                + ", " + "N'" + data[6] + "'"            // 9.画面・機能名
                                + ", " + "N'" + data[7] + "'"            // 10.帳票分類名
                                + ", " + "N'" + data[8] + "'"            // 11.帳票名
                                + ", " + "N'" + data[9] + "'"            // 12.雛型ファイル名
                                + ", " + "N'" + data[10] + "'"           // 13.ダウンロードファイル名
                                + ")"
                                ;

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

        // Garoon送信ボタンの更新処理
        public void garoonRenkeiUpdate(String MadoguchiID,String[] UserInfos, out string mes)
        {
            string methodName = ".garoonRenkeiUpdate";
            mes = "";
            // Garoon連携フラグ  1:連携 0:チェックなし
            string GaroonRenkei = "0";

            DataTable combodt = new DataTable();
            combodt = getData("MadoguchiGaroonRenkei", "MadoguchiGaroonRenkei", "MadoguchiJouhou", "MadoguchiID = '" + MadoguchiID + "' ");

            if (combodt != null && combodt.Rows.Count > 0)
            {
                GaroonRenkei = combodt.Rows[0][0].ToString();
            }

            // エラーフラグ true:エラー false:正常
            Boolean errorFlg = false;
            // Garoon連携対象チェック
            if (GaroonRenkei == "0")
            {
                // I20005:Garoonとの連携対象ではありません。
                mes += GetMessage("I20005", "");
                errorFlg = true;
            }

            string MadoguchiTantoushaCD = "";
            combodt = new DataTable();
            combodt = getData("MadoguchiTantoushaCD", "MadoguchiTantoushaCD", "MadoguchiJouhou", "MadoguchiID = '" + MadoguchiID + "' ");

            if (combodt != null && combodt.Rows.Count > 0)
            {
                MadoguchiTantoushaCD = combodt.Rows[0][0].ToString();
            }

            // 窓口担当者チェック
            //if (item1_MadoguchiTantousha.Text == "")
            if (MadoguchiTantoushaCD == "")
            {
                // E20011:窓口担当者が未登録のため、Garoon連携ができません。
                mes += GetMessage("E20011", "");
                errorFlg = true;
            }

            // 連携処理
            if (errorFlg == false)
            {
                string w_MadoguchiMailGaRenkeiKubun = "";
                string w_MadoguchiMailMessageID = "";
                string w_MadoguchiUketsukeBangou = "";
                string w_MadoguchiUketsukeBangouEdaban = "";
                string w_MadoguchiTantoushaCD = "";
                string w_TokuchoBangou = "";
                string w_MailInfoCSVWorkAtesakiUser = "";
                //string w_KojinCD = "";
                string w_MadoguchiKanriGijutsusha = "";
                string w_MadoguchiL1ChousaBushoCD = "";
                string w_MadoguchiL1ChousaTantoushaCD = "";

                string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();
                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;
                    try
                    {
                        string historyMessage = "";

                        // I20006:Garoon送信ボタンからOKが押下されました。
                        historyMessage = GetMessage("I20006", "") + " ID:" + MadoguchiID + " Garoon連携区分:" + GaroonRenkei;

                        //履歴登録
                        //cmd.CommandText = "INSERT INTO T_HISTORY(" +
                        //    "H_DATE_KEY " +
                        //    ",H_NO_KEY " +
                        //    ",H_OPERATE_DT " +
                        //    ",H_OPERATE_USER_ID " +
                        //    ",H_OPERATE_USER_MEI " +
                        //    ",H_OPERATE_USER_BUSHO_CD " +
                        //    ",H_OPERATE_USER_BUSHO_MEI " +
                        //    ",H_OPERATE_NAIYO " +
                        //    ",H_ProgramName " +
                        //    ",MadoguchiID " +
                        //    ",HistoryBeforeTantoubushoCD " +
                        //    ",HistoryBeforeTantoushaCD " +
                        //    ",HistoryAfterTantoubushoCD " +
                        //    ",HistoryAfterTantoushaCD " +
                        //    ")VALUES(" +
                        //    "SYSDATETIME() " +
                        //    ", " + getSaiban("HistoryID") + " " +
                        //    ",SYSDATETIME() " +
                        //    ",'" + UserInfos[0] + "' " +
                        //    ",'" + UserInfos[1] + "' " +
                        //    ",'" + UserInfos[2] + "' " +
                        //    ",'" + UserInfos[3] + "' " +
                        //    ",'" + historyMessage + "'" +
                        //    ",'" + pgmName + methodName + "' " +
                        //    "," + MadoguchiID + " " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ")";

                        //cmd.ExecuteNonQuery();

                        Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], historyMessage, pgmName + methodName, MadoguchiID);

                        var Dt = new System.Data.DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "MadoguchiGaroonRenkei,MadoguchiUketsukeBangou,MadoguchiUketsukeBangouEdaban,MadoguchiTantoushaCD,MadoguchiKanriGijutsusha " +
                          "FROM MadoguchiJouhou " +
                          "WHERE MadoguchiID = '" + MadoguchiID + "' ";

                        //データ取得
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(Dt);

                        // MadoguchiJouhouに登録されているデータを取得
                        if (Dt != null && Dt.Rows.Count > 0)
                        {
                            w_MadoguchiMailGaRenkeiKubun = Dt.Rows[0][0].ToString();
                            w_MadoguchiUketsukeBangou = Dt.Rows[0][1].ToString();
                            w_MadoguchiUketsukeBangouEdaban = Dt.Rows[0][2].ToString();
                            w_MadoguchiTantoushaCD = Dt.Rows[0][3].ToString();
                            w_MadoguchiKanriGijutsusha = Dt.Rows[0][4].ToString();
                        }

                        // MadoguchiMailのIDを取得
                        string discript = "MadoguchiMailMessageID ";
                        string value = "TOP 1 MadoguchiMailMessageID ";
                        string table = "MadoguchiMail ";
                        string where = "MadoguchiMailTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + w_MadoguchiUketsukeBangou + "' AND MadoguchiMailTokuchoBangouEda COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + w_MadoguchiUketsukeBangouEdaban + "' ";

                        // データ取得
                        DataTable tmpdt = getData(discript, value, table, where);
                        if (tmpdt != null && tmpdt.Rows.Count > 0)
                        {
                            w_MadoguchiMailMessageID = tmpdt.Rows[0][0].ToString();
                        }
                        else
                        {
                            // 取得できなかった場合は、0をセット（MailInfoCSVWorkMessageID は数値型の為、insert時に空文字をセットしようとするとエラーになる）
                            w_MadoguchiMailMessageID = "0";
                        }

                        // 特調番号
                        w_TokuchoBangou = w_MadoguchiUketsukeBangou + "-" + w_MadoguchiUketsukeBangouEdaban;

                        // 調査員取得
                        w_MailInfoCSVWorkAtesakiUser = SetChousain(w_MadoguchiTantoushaCD, w_MailInfoCSVWorkAtesakiUser, MadoguchiID);

                        // 管理技術者が存在すれば
                        if(w_MadoguchiKanriGijutsusha != "" && w_MadoguchiKanriGijutsusha != "0")
                        {
                            w_MailInfoCSVWorkAtesakiUser = SetChousain(w_MadoguchiKanriGijutsusha, w_MailInfoCSVWorkAtesakiUser, MadoguchiID);
                        }

                        Dt = new System.Data.DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "MadoguchiL1ChousaBushoCD,MadoguchiL1ChousaTantoushaCD " +
                          "FROM MadoguchiJouhouMadoguchiL1Chou " +
                          "WHERE MadoguchiID = '" + MadoguchiID + "' Order By MadoguchiL1ChousaCD";

                        //データ取得
                        sda = new SqlDataAdapter(cmd);
                        sda.Fill(Dt);

                        // MadoguchiJouhouに登録されているデータを取得
                        if (Dt != null && Dt.Rows.Count > 0)
                        {
                            for (int i = 0; i < Dt.Rows.Count; i++)
                            {
                                //w_MadoguchiL1ChousaBushoCD = Dt.Rows[0][0].ToString();
                                w_MadoguchiL1ChousaTantoushaCD = Dt.Rows[i][1].ToString();

                                // 調査担当者が存在すれば
                                if (w_MadoguchiL1ChousaTantoushaCD != "" && w_MadoguchiL1ChousaTantoushaCD != "0")
                                {
                                    w_MailInfoCSVWorkAtesakiUser = SetChousain(w_MadoguchiL1ChousaTantoushaCD, w_MailInfoCSVWorkAtesakiUser, MadoguchiID);
                                }
                                // 支部応援マスタから担当調査員の部所に該当する調査員を設定する
                                if (w_MadoguchiL1ChousaBushoCD != Dt.Rows[i][0].ToString())
                                {
                                    w_MadoguchiL1ChousaBushoCD = Dt.Rows[i][0].ToString();
                                    w_MailInfoCSVWorkAtesakiUser = GetShibuouen(w_MadoguchiL1ChousaBushoCD, w_MailInfoCSVWorkAtesakiUser);
                                }
                            }
                        }

                        // Garoon追加追加宛先の調査員も追加する
                        DataTable GaroonDt = new System.Data.DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "GaroonTsuikaAtesakiBushoCD,GaroonTsuikaAtesakiTantoushaCD " +
                          "FROM GaroonTsuikaAtesaki " +
                          "WHERE GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "' " +
                          " AND GaroonTsuikaAtesakiDeleteFlag <> 1";

                        //データ取得
                        sda = new SqlDataAdapter(cmd);
                        sda.Fill(GaroonDt);

                        // GaroonTsuikaAtesakiに登録されているデータを取得
                        if (GaroonDt != null && GaroonDt.Rows.Count > 0)
                        {
                            for (int i = 0; i < GaroonDt.Rows.Count; i++)
                            {
                                //w_MadoguchiL1ChousaBushoCD = Dt.Rows[0][0].ToString();
                                w_MadoguchiL1ChousaTantoushaCD = GaroonDt.Rows[i][1].ToString();

                                // 調査担当者が存在すれば
                                if (w_MadoguchiL1ChousaTantoushaCD != "" && w_MadoguchiL1ChousaTantoushaCD != "0")
                                {
                                    w_MailInfoCSVWorkAtesakiUser = SetChousain(w_MadoguchiL1ChousaTantoushaCD, w_MailInfoCSVWorkAtesakiUser, MadoguchiID);
                                }
                            }
                        }

                        // メール情報CSVに追加するユーザーが空でない場合
                        if (w_MailInfoCSVWorkAtesakiUser != "")
                        {
                            // メール情報ワークの取得
                            Dt = new System.Data.DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "MailInfoCSVWorkID " +
                              "FROM MailInfoCSVWork " +
                              "WHERE MailInfoCSVWorkTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + w_TokuchoBangou + "' AND MailInfoCSVWorkCSVOutFlg = 0 AND MailInfoCSVWorkGaRenkeiFlg = 0 AND MailInfoCSVWorkDeleteFlag = 0";

                            //データ取得
                            sda = new SqlDataAdapter(cmd);
                            sda.Fill(Dt);

                            string MailInfoCSVWorkID = "";
                            // データの存在確認
                            if (Dt != null && Dt.Rows.Count > 0)
                            {
                                MailInfoCSVWorkID = Dt.Rows[0][0].ToString();
                                // 連携フラグにより、更新、削除を振り分ける
                                if (w_MadoguchiMailGaRenkeiKubun == "1")
                                {
                                    // 宛先を更新
                                    cmd.CommandText = "UPDATE MailInfoCSVWork SET " +
                                    "MailInfoCSVWorkAtesakiUser = '" + w_MailInfoCSVWorkAtesakiUser + "' " +
                                    ",MailInfoCSVWorkUpdateDate = SYSDATETIME() " +
                                    ",MailInfoCSVWorkUpdateUser = N'" + UserInfos[0] + "' " +
                                    ",MailInfoCSVWorkUpdateProgram = '" + pgmName + methodName + "' " +
                                    "Where MailInfoCSVWorkID = '" + MailInfoCSVWorkID + "' " +
                                    "AND MailInfoCSVWorkDeleteFlag = 0";
                                    cmd.ExecuteNonQuery();
                                }
                                else
                                {
                                    // 連携フラグがないので削除
                                    cmd.CommandText = "DELETE FROM MailInfoCSVWork " +
                                        "WHERE MailInfoCSVWorkID = '" + MailInfoCSVWorkID + "' ";
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                if (w_MadoguchiMailGaRenkeiKubun == "1")
                                {
                                    // 存在しない場合、Insert
                                    cmd.CommandText = "INSERT INTO MailInfoCSVWork (" +
                                    "MailInfoCSVWorkID " +
                                    ",MailInfoCSVWorkMadoguchiID " +
                                    ",MailInfoCSVWorkTokuchoBangou " +
                                    ",MailInfoCSVWorkMessageID " +
                                    ",MailInfoCSVWorkAtesakiUser " +
                                    ",MailInfoCSVWorkCSVOutFlg " +
                                    ",MailInfoCSVWorkGaRenkeiFlg " +
                                    ",MailInfoCSVWorkCreateDate " +
                                    ",MailInfoCSVWorkCreateUser " +
                                    ",MailInfoCSVWorkCreateProgram " +
                                    ",MailInfoCSVWorkUpdateDate " +
                                    ",MailInfoCSVWorkUpdateUser " +
                                    ",MailInfoCSVWorkUpdateProgram " +
                                    ",MailInfoCSVWorkDeleteFlag " +
                                    ") VALUES (" +
                                    getSaiban("MailInfoCSVWorkID") +
                                    ",'" + MadoguchiID + "' " +
                                    ",N'" + w_TokuchoBangou + "' " +
                                    ",'" + w_MadoguchiMailMessageID + "' " +
                                    ",'" + w_MailInfoCSVWorkAtesakiUser + "' " +
                                    ",'0'" +
                                    ",'0'" +
                                    ",SYSDATETIME()" +
                                    ",N'" + UserInfos[0] + "' " +
                                    ",'" + pgmName + methodName + "'" +
                                    ",SYSDATETIME()" +
                                    ",N'" + UserInfos[0] + "' " +
                                    ",'" + pgmName + methodName + "'" +
                                    ",0" + 
                                    ")";
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                        else
                        {
                            // 宛先が存在しない場合、削除する
                            cmd.CommandText = "DELETE FROM MailInfoCSVWork " +
                                             "WHERE MailInfoCSVWorkTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + w_TokuchoBangou + "' AND MailInfoCSVWorkCSVOutFlg = 0 AND MailInfoCSVWorkGaRenkeiFlg = 0";
                            cmd.ExecuteNonQuery();
                        }

                        // 窓口情報の連携実行日時を更新
                        string datetTime = DateTime.Now.ToString();

                        cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                        "MadoguchiGaroonRenkeiJikouDate = '" + datetTime + "' " +
                        "Where MadoguchiID = '" + MadoguchiID + "' ";
                        cmd.ExecuteNonQuery();

                        transaction.Commit();

                        outputLogger("GaroonBtn_Click", GetMessage("I20006", "") + " ID:" + MadoguchiID + " Garoon連携区分:1", "insert", "DEBUG");
                        // I20004:Garoonとの連携を行いました。
                        mes += GetMessage("I20004", "");

                        conn.Close();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                    finally
                    {
                        conn.Close();
                    }
                    cmd.Transaction = transaction;
                }
            }
        }

        // 個人コードを検索し、連結して返す
        private string SetChousain(string KojinCD, string MailInfoCSVWorkAtesakiUser, string MadoguchiID)
        {
            string discript = "KojinCD ";
            string value = "KojinCD ";
            string table = "Mst_Chousain ";
            string where = "KojinCD = '" + KojinCD + "' ";
            string w_KojinCD = "";

            // データ取得
            var tmpdt = getData(discript, value, table, where);
            if (tmpdt != null && tmpdt.Rows.Count > 0)
            {
                if (MailInfoCSVWorkAtesakiUser == "")
                {
                    MailInfoCSVWorkAtesakiUser = tmpdt.Rows[0][0].ToString();
                }
                else
                {
                    w_KojinCD = tmpdt.Rows[0][0].ToString();
                    // MailInfoCSVWorkAtesakiUser が2048文字までなので、OVERする場合は、セットしない
                    if ((MailInfoCSVWorkAtesakiUser.Length + w_KojinCD.Length) <= 2048)
                    {
                        // 既に存在する場合は追加しない
                        if(MailInfoCSVWorkAtesakiUser.IndexOf(w_KojinCD) == -1)
                        {
                            MailInfoCSVWorkAtesakiUser = MailInfoCSVWorkAtesakiUser + ";" + w_KojinCD;
                        }
                    }
                    else
                    {
                        outputLogger("SetChousain", "ID:" + MadoguchiID + " Garoon連携で宛先ユーザーの文字数が2048を超える為、KojinCD:" + w_KojinCD + " を追加できませんでした。", "insert", "DEBUG");
                    }
                }
            }
            return MailInfoCSVWorkAtesakiUser;
        }

        // 支部応援の取得
        private string GetShibuouen(string w_MadoguchiL1ChousaBushoCD, string MailInfoCSVWorkAtesakiUser)
        {
            string discript = "ShibuouenKojinCD ";
            string value = "ShibuouenKojinCD ";
            string table = "Mst_Shibuouen ";
            //string where = "ShibuouenDeleteFlag = 0 Order By ShibuouenKojinCD ";
            string where = "(ShibuouenDeleteFlag = 0 or ShibuouenDeleteFlag = 1) Order By ShibuouenKojinCD ";
            string w_ShibuouenKojinCD = "";

            // データ取得
            var tmpdt = getData(discript, value, table, where);
            DataTable dt = new DataTable();

            if (tmpdt != null && tmpdt.Rows.Count > 0)
            {
                for (int i = 0; i < tmpdt.Rows.Count; i++)
                {
                    w_ShibuouenKojinCD = tmpdt.Rows[i][0].ToString();

                    discript = "KojinCD ";
                    value = "KojinCD ";
                    table = "Mst_Chousain ";
                    where = "GyoumuBushoCD = '" + w_MadoguchiL1ChousaBushoCD + "' AND KojinCD = '" + w_ShibuouenKojinCD + "' ";
                    dt = getData(discript, value, table, where);

                    // 存在する場合のみ追加する
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        if (MailInfoCSVWorkAtesakiUser == "")
                        {
                            MailInfoCSVWorkAtesakiUser = dt.Rows[0][0].ToString();
                        }
                        else
                        {
                            // 既に存在する場合は追加しない
                            if (MailInfoCSVWorkAtesakiUser.IndexOf(dt.Rows[0][0].ToString()) == -1)
                            {
                                MailInfoCSVWorkAtesakiUser = MailInfoCSVWorkAtesakiUser + ";" + dt.Rows[0][0].ToString();
                            }
                        }
                    }
                }
            }
            return MailInfoCSVWorkAtesakiUser;
        }

        public class CTabPage : System.Windows.Forms.TabPage
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

        // 皇帝まもる連携
        public void KouteiTantouBushoRenkei(string MadoguchiID, string KojinCD, string BushoCD)
        {
            string Pgmname = "KouteiTantouBushoRenkei";
            //Processオブジェクトを作成
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            //ComSpec(cmd.exe)のパスを取得して、FileNameプロパティに指定
            p.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec");

            //出力を読み取れるようにする
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardInput = false;
            //ウィンドウを表示しないようにする
            p.StartInfo.CreateNoWindow = true;

            // GeneXusのexeは共有フォルダに配置する apromainkouteitantou.exe
            p.StartInfo.Arguments = @"/c " + GetCommonValue1("KOUTEI_TANTOU_EXE_FOLDER") + " " + MadoguchiID + " " + Pgmname + " " + KojinCD + " " + BushoCD;

            //起動
            p.Start();

            //出力を読み取る
            string results = p.StandardOutput.ReadToEnd();

            //プロセス終了まで待機する
            //WaitForExitはReadToEndの後である必要がある
            //(親プロセス、子プロセスでブロック防止のため)
            p.WaitForExit();
            p.Close();

            string[] result = results.Replace(Environment.NewLine, "").Split('|');

            // result
            // 成否判定 0:正常 1：エラー
            if (result != null && result.Length >= 1)
            {
                // 改行コードがあるので、削る
                result[0] = result[0].Replace(@"\r\n", "");

                if (result[0].Trim() == "1")
                {
                    // エラーが発生
                    outputLogger(Pgmname, "皇帝まもるとの連携に失敗しました。", "insert", "DEBUG");

                }
                else if (result[0].Trim() == "0")
                {
                    // 正常

                }
            }
            else
            {
                // エラーが発生しました
                outputLogger(Pgmname, "皇帝まもるexeファイルの呼び出しに失敗しました。", "insert", "DEBUG");
            }
        }

        /// <summary>
        /// 文字列の指定した位置から指定した長さを取得する
        /// </summary>
        /// <param name="str">文字列</param>
        /// <param name="start">開始位置</param>
        /// <param name="len">長さ</param>
        /// <returns>取得した文字列</returns>
        public static string Mid(string str, int start, int len)
        {
            if (start <= 0)
            {
                throw new ArgumentException("引数'start'は1以上でなければなりません。");
            }
            if (len < 0)
            {
                throw new ArgumentException("引数'len'は0以上でなければなりません。");
            }
            if (str == null || str.Length < start)
            {
                return "";
            }
            if (str.Length < (start + len))
            {
                return str.Substring(start - 1);
            }
            return str.Substring(start - 1, len);
        }

        /// <summary>
        /// 文字列の指定した位置から末尾までを取得する
        /// </summary>
        /// <param name="str">文字列</param>
        /// <param name="start">開始位置</param>
        /// <returns>取得した文字列</returns>
        public string Mid(string str, int start)
        {
            return Mid(str, start, str.Length);
        }

        /// <summary>
        /// 文字列の先頭から指定した長さの文字列を取得する
        /// </summary>
        /// <param name="str">文字列</param>
        /// <param name="len">長さ</param>
        /// <returns>取得した文字列</returns>
        public string Left(string str, int len)
        {
            if (len < 0)
            {
                throw new ArgumentException("引数'len'は0以上でなければなりません。");
            }
            if (str == null)
            {
                return "";
            }
            if (str.Length <= len)
            {
                return str;
            }
            return str.Substring(0, len);
        }

        /// <summary>
        /// 文字列の末尾から指定した長さの文字列を取得する
        /// </summary>
        /// <param name="str">文字列</param>
        /// <param name="len">長さ</param>
        /// <returns>取得した文字列</returns>
        public string Right(string str, int len)
        {
            if (len < 0)
            {
                throw new ArgumentException("引数'len'は0以上でなければなりません。");
            }
            if (str == null)
            {
                return "";
            }
            if (str.Length <= len)
            {
                return str;
            }
            return str.Substring(str.Length - len, len);
        }

        //不具合管理表No1017（751）
        //タブの表示を加工する。タブコントロールのDrawModeプロパティを下記のように設定する必要あり
        // tab.DrawMode = TabDrawMode.OwnerDrawFixed;
        public void tabDisplaySet(TabControl tab ,object sender, DrawItemEventArgs e)
        {
            System.Drawing.SolidBrush backBrush;
            System.Drawing.SolidBrush foreBrush;
            System.Drawing.Font font;
            if (tab.SelectedIndex == e.Index)
            {
                backBrush = new System.Drawing.SolidBrush(System.Drawing.SystemColors.Window);
                foreBrush = new System.Drawing.SolidBrush(System.Drawing.Color.Blue);
                font = new System.Drawing.Font("メイリオ", 14, System.Drawing.FontStyle.Regular);
            }
            else
            {
                backBrush = new System.Drawing.SolidBrush(System.Drawing.SystemColors.Control);
                foreBrush = new System.Drawing.SolidBrush(System.Drawing.Color.Black);
                font = new System.Drawing.Font("メイリオ", 14, System.Drawing.FontStyle.Regular);
            }
            System.Drawing.StringFormat format = new System.Drawing.StringFormat();
            System.Drawing.RectangleF rect = new System.Drawing.RectangleF(e.Bounds.X, e.Bounds.Y + 6, e.Bounds.Width, e.Bounds.Height);
            format.Alignment = System.Drawing.StringAlignment.Center;
            e.Graphics.FillRectangle(backBrush, e.Bounds);
            e.Graphics.DrawString(tab.TabPages[e.Index].Text, font, foreBrush, rect, format);
        }

        //不具合管理表No1228（919） ファイルが開いていてロックされてないか
        /// 0:ファイルが編集可能／1:ファイルが存在しない／2:ファイルが開いていてロックされてる
        public int getFileStatus(string path)
        {
            FileStream stream = null;
            if (!File.Exists(path))
            {
                return 1;
            }
            try
            {
                stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch
            {
                return 2;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
            }

            return 0;
        }
    }
}
