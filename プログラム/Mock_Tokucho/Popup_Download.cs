using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace TokuchoBugyoK2
{
    public partial class Popup_Download : Form
    {
        public string TotalFilePath = null;
        public string ExcelName = null;
        public string ExcelPath = null;

        public Popup_Download()
        {
            InitializeComponent();
        }

        // 開くボタン
        private void button5_Click(object sender, EventArgs e)
        {
            GlobalMethod GlobalMethod = new GlobalMethod();
            // そのまま開くとWORKの名前が見えてしまうので、
            // フォルダを作成し、その中にリネームしたファイルをコピーし、開く
            string path = "";
            string fileName = "";
            string fileRenban = "";
            string folderPath = "";

            //// ファイルパスに / が含まれる可能性を考慮し、念の為、/ を \ に置換しておく
            //path = TotalFilePath.Replace("/", @"\");
            //// ファイル名の前の \ の位置を取得
            //int pos = path.LastIndexOf(@"\");

            try
            {
                // パスからファイル名を取得
                //fileName = path.Substring(pos + 1, path.Length - pos - 1);
                fileName = Path.GetFileName(TotalFilePath);

                if (fileName.Length > 14) {
                    // C#で出力したファイルと、GeneXusを通して出力した場合を切り分ける
                    if (fileName.Substring(0,5) != "年度計画表") {
                        // 頭から14桁の連番部分を取得する 例：00000000394512_エントリーシート.xlsx
                        fileRenban = fileName.Substring(0, 14);
                        // フォルダパス部分のみを取得する
                        folderPath = TotalFilePath.Replace(fileName, "");
                        // フォルダパス + 連番部分でフォルダを作成する
                        folderPath = folderPath + fileRenban;
                        DirectoryInfo di = new DirectoryInfo(folderPath);
                        di.Create();

                        // 作成した連番フォルダにファイルをコピーする
                        System.IO.File.Copy(TotalFilePath, folderPath + @"\" + ExcelName, true);

                        string Folderpath = System.IO.Path.Combine(folderPath + @"\" + ExcelName, "");
                        System.Diagnostics.Process.Start(folderPath + @"\" + ExcelName);

                        this.Close();
                    }
                    else
                    {
                        // C#側で出力した場合は、そのままファイルを開く
                        string Folderpath = System.IO.Path.Combine(TotalFilePath, "");
                        System.Diagnostics.Process.Start(Folderpath);
                    }
                }
            }
            catch
            {
                MessageBox.Show("EXCELファイルを開く処理に失敗しました。");
                // エラー
                GlobalMethod.outputLogger("PopUp_DownLoad", "ファイルを開くエラー:" + TotalFilePath, "", "DEBUG");
            }
        }

        // 保存ボタン
        private void button1_Click(object sender, EventArgs e)
        {
            //SaveFileDialogを生成する
            SaveFileDialog sa = new SaveFileDialog();
            sa.Title = "ファイルを保存する";
            sa.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            sa.FileName = ExcelName;
            sa.Filter = @"Excel ファイル|*.xls;*.xlsx;*.xlsm|全てのファイル|*.*";
            //sa.Filter = "テキストファイル(*.txt;*.text)|*.txt;*.text|すべてのファイル(*.*)|*.*";
            sa.FilterIndex = 1;


            //オープンファイルダイアログを表示する
            DialogResult result = sa.ShowDialog();

            if (result == DialogResult.OK)
            {
                //「保存」ボタンが押された時の処理
                string fileName = System.IO.Path.Combine(TotalFilePath, "");
                try
                {
                    System.IO.File.Copy(fileName, sa.FileName, true);
                }
                catch (Exception)
                {
                    MessageBox.Show("EXCELファイルの保存に失敗しました。");
                }
                ///System.Diagnostics.Process.Start(sa.FileName);

                this.Close();
            }
            else if (result == DialogResult.Cancel)
            {
                //「キャンセル」ボタンまたは「×」ボタンが選択された時の処理
            }
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Popup_Download_Load(object sender, EventArgs e)
        {
            this.Location = new System.Drawing.Point(10, this.Parent.Height - this.Height - 30);
            //フォームが最大化されないようにする
            this.MaximizeBox = false;
            //フォームが最小化されないようにする
            this.MinimizeBox = false;
            // 拡大縮小禁止
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            // タイトルバーを消す
            //this.FormBorderStyle = FormBorderStyle.None;

            if (TotalFilePath == null)
            {
                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
                TotalFilePath = System.IO.Path.Combine(stCurrentDir, ExcelPath);
            }
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_SYSCOMMAND = 0x0112;
            const long SC_MOVE = 0xF010L;

            if (m.Msg == WM_SYSCOMMAND &&
                (m.WParam.ToInt64() & 0xFFF0L) == SC_MOVE)
            {
                m.Result = IntPtr.Zero;
                return;
            }

            base.WndProc(ref m);
        }
    }
}
