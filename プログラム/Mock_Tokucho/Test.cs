using C1.C1Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TokuchoBugyoK2
{
    public partial class Test : Form
    {
        public Test()
        {
            InitializeComponent();
        }

        private void 貼り付けToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IDataObject data = Clipboard.GetDataObject();
            if (data.GetDataPresent(DataFormats.Text))
            {
                string str1, str2;
                str1 = (string)data.GetData(DataFormats.Text);
                str2 = str1.Remove(str1.Length - 1, 1);
                /*
                c1FlexGrid1.Select(c1FlexGrid1.Row, c1FlexGrid1.Col, c1FlexGrid1.Rows.Count - 1, c1FlexGrid1.Cols.Count - 1);
                c1FlexGrid1.Clip = str2;*/
            }
        }

        private void c1FlexGrid1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.V && e.Control == true)
            {
                IDataObject data = Clipboard.GetDataObject();
                if (data.GetDataPresent(DataFormats.Text))
                {
                    string str1, str2;
                    str1 = (string)data.GetData(DataFormats.Text);
                    str2 = str1.Remove(str1.Length - 1, 1);
                    /*
                    c1FlexGrid1.Select(c1FlexGrid1.Row, c1FlexGrid1.Col, c1FlexGrid1.Rows.Count - 1, c1FlexGrid1.Cols.Count - 1);
                    c1FlexGrid1.Clip = str2;*/
                }
            }
        }

        private void Test_Load(object sender, EventArgs e)
        {



            //Excel・シート取込
            string[] fileName = { @"test1.xlsx", @"test2.xlsx", @"test3.xlsx", @"test4.xlsx", @"test5.xlsm", @"test6.xlsm" };

            for (int i = 0; i < fileName.Length; i++)
            {
                C1XLBook wb = getExcelFile(@"C:\Users\mohara.tomoaki\Downloads\新しいフォルダー\調査\取込み前\" + fileName[i]);

                //testファイル用出力データ取得・セット処理
                setTestExcelFile(wb);

                //Excel出力
                wb.Save(@"C:\Users\mohara.tomoaki\Downloads\新しいフォルダー\調査\取込み後\" + fileName[i]);
            }

            System.Web.Mail.MailMessage mm = new System.Web.Mail.MailMessage();
            //送信者
            mm.From = "sender@xxx.xx.com";
            //送信先
            mm.To = "recipient1@xxx.xx.com";
            //題名
            mm.Subject = "テスト";
            //本文
            mm.Body = "こんにちは。これはテストです。";
            //JISコードに変換する
            mm.BodyEncoding = System.Text.Encoding.GetEncoding(50220);
            //SMTPサーバーを指定する
            System.Web.Mail.SmtpMail.SmtpServer = "（SMTPサーバーを指定する）";
            //送信する
            System.Web.Mail.SmtpMail.Send(mm);
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(textBox23.Text);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox2.Text = Microsoft.VisualBasic.Strings.StrConv(textBox1.Text, Microsoft.VisualBasic.VbStrConv.Narrow, 0x411);
        }

        private C1XLBook getExcelFile(String fileName)
        {
            //Excel取込
            C1XLBook wb = new C1XLBook();
            if (!System.IO.File.Exists(fileName))
            {
                MessageBox.Show("'" + fileName + "'は存在しません。");
                return null;
            }

            wb.Load(fileName);
            return wb;
        }


        private C1XLBook setTestExcelFile(C1.C1Excel.C1XLBook wb)
        {
            //シート取込
            int u = wb.Sheets.Count;
            XLSheet sheet1 = wb.Sheets[u - 1];

            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {

                var cmd = conn.CreateCommand();
                //部所データ取得
                var comboDt = new DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                "GyoumuBushoCD ,ShibuMei ,KaMei " +
                "FROM " + "Mst_Busho ";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(comboDt);

                //データセット
                for (int i = 0; comboDt.Rows.Count > i; i++)
                {
                    DataRow nr = comboDt.Rows[i];
                    sheet1[i + 1, 0].Value = nr["GyoumuBushoCD"].ToString();
                    sheet1[i + 1, 1].Value = nr["ShibuMei"].ToString();
                    sheet1[i + 1, 2].Value = nr["KaMei"].ToString();

                }

                //調査員データ取得
                var comboDt2 = new DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                "KojinCD ,ChousainMei " +
                "FROM " + "Mst_Chousain ";

                //データ取得
                sda = new SqlDataAdapter(cmd);
                sda.Fill(comboDt2);

                //データセット
                for (int i = 0; comboDt2.Rows.Count > i; i++)
                {
                    DataRow nr = comboDt2.Rows[i];
                    sheet1[i + 1, 4].Value = nr["KojinCD"].ToString();
                    sheet1[i + 1, 5].Value = nr["ChousainMei"].ToString();

                }
            }
            return wb;
        }

    }
}
